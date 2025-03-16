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
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Classroom Assignment Downloader</title>
      <style>
        body {
          font-family: 'Google Sans', Roboto, Arial, sans-serif;
          line-height: 1.6;
          color: #202124;
          max-width: 800px;
          margin: 0 auto;
          padding: 20px;
          background-color: #f8f9fa;
        }
        
        h2 {
          color: #1a73e8;
          margin-bottom: 24px;
          padding-bottom: 8px;
          border-bottom: 1px solid #dadce0;
        }
        
        h3 {
          color: #202124;
          margin-top: 24px;
          margin-bottom: 12px;
          font-weight: 500;
        }
        
        select, button {
          font-family: 'Google Sans', Roboto, Arial, sans-serif;
          font-size: 14px;
          border-radius: 4px;
          border: 1px solid #dadce0;
          padding: 8px 16px;
          background-color: white;
          transition: all 0.2s;
        }
        
        select {
          width: 100%;
          margin: 10px 0;
          height: 40px;
          cursor: pointer;
        }
        
        select:focus {
          outline: none;
          border-color: #1a73e8;
          box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.2);
        }
        
        button {
          cursor: pointer;
          font-weight: 500;
          min-height: 40px;
        }
        
        button:hover:not(:disabled) {
          background-color: #f1f3f4;
        }
        
        button:focus {
          outline: none;
          box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.2);
        }
        
        button:disabled {
          opacity: 0.6;
          cursor: not-allowed;
        }
        
        #downloadBtn {
          background-color: #1a73e8;
          color: white;
        }
        
        #downloadBtn:hover:not(:disabled) {
          background-color: #1765cc;
        }
        
        #cancelBtn {
          background-color: #ea4335;
          color: white;
        }
        
        #cancelBtn:hover:not(:disabled) {
          background-color: #d93025;
        }
        
        .button-container {
          display: flex;
          gap: 12px;
          margin-top: 24px;
        }
        
        .topic-item {
          margin: 8px 0;
          padding: 8px;
          border-radius: 4px;
          transition: background-color 0.2s;
        }
        
        .topic-item:hover {
          background-color: #f1f3f4;
        }
        
        .topic-item label {
          display: flex;
          align-items: center;
          cursor: pointer;
        }
        
        .topic-item input {
          margin-right: 10px;
        }
        
        #selectAllContainer {
          margin-top: 16px;
          padding: 8px;
          border-top: 1px solid #dadce0;
          display: flex;
          align-items: center;
        }
        
        #selectAllContainer input {
          margin-right: 10px;
        }
        
        #status {
          margin-top: 24px;
          color: #5f6368;
          min-height: 24px;
        }
        
        #folderLink {
          margin-top: 12px;
        }
        
        #folderLink a {
          color: #1a73e8;
          text-decoration: none;
          display: inline-flex;
          align-items: center;
        }
        
        #folderLink a:hover {
          text-decoration: underline;
        }
        
        #folderLink a::before {
          content: '';
          display: inline-block;
          width: 18px;
          height: 18px;
          margin-right: 8px;
          background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="%231a73e8"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14z"/><path d="M7 12h10v2H7z"/><path d="M7 7h10v2H7z"/><path d="M7 17h7v2H7z"/></svg>');
          background-size: contain;
        }
        
        .progress-container {
          margin-top: 16px;
          background-color: #e8eaed;
          border-radius: 4px;
          overflow: hidden;
        }
        
        .progress-bar {
          height: 4px;
          background-color: #1a73e8;
          width: 0%;
          transition: width 0.3s ease;
        }
        
        .loading-spinner {
          border: 3px solid rgba(26, 115, 232, 0.2);
          border-top: 3px solid #1a73e8;
          border-radius: 50%;
          width: 24px;
          height: 24px;
          animation: spin 1.5s linear infinite;
          margin: 20px auto;
          display: inline-block;
        }
        
        .inline-spinner {
          width: 16px;
          height: 16px;
          margin-right: 8px;
          vertical-align: middle;
        }
        
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        .card {
          background-color: white;
          border-radius: 8px;
          box-shadow: 0 1px 2px 0 rgba(60, 64, 67, 0.3), 0 1px 3px 1px rgba(60, 64, 67, 0.15);
          padding: 24px;
          margin-bottom: 24px;
        }
        
        .option-item {
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 8px;
          cursor: pointer;
        }
        
        .option-item:hover {
          background-color: #f1f3f4;
          border-radius: 4px;
        }
      </style>
    </head>
    <body>
    <h2>Google Classroom Assignment Downloader</h2>
      
      <div class="card">
        <h3>Step 1: Select Course</h3>
        <select id="courseSelect" onchange="loadTopics()">
          <option value="">Loading courses...</option>
        </select>
      </div>
      
      <div id="topicsContainer" class="card" style="display:none;">
        <h3>Step 2: Select Topics</h3>
        <div id="topicsList">
          <div class="loading-spinner"></div>
          </div>
        <div id="selectAllContainer" style="display:none;">
          <input type="checkbox" id="selectAll" onclick="toggleAllTopics()">
          <label for="selectAll">Select All Topics</label>
        </div>
      </div>
      
      <div class="card">
        <h3>Options</h3>
        <label class="option-item">
          <input type="checkbox" id="convertToPdf">
          Convert compatible files to PDF
        </label>
      </div>
      
      <div class="button-container">
        <button id="downloadBtn" onclick="downloadSelected()" disabled>
          Download Selected Assignments
        </button>
        <button id="cancelBtn" onclick="cancelDownload()" style="display:none;">
          Cancel Download
        </button>
      </div>
      
      <div id="status"></div>
      <div id="folderLink" style="display:none;"></div>
      <div id="progress" style="display:none;" class="progress-container">
        <div id="progressBar" class="progress-bar"></div>
      </div>
    
    <script>
        let downloadInProgress = false;
        let folderUrl = '';
        
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

          document.getElementById('status').innerHTML = '<div class="loading-spinner inline-spinner"></div> Loading topics...';
          document.getElementById('topicsContainer').style.display = 'block';
          document.getElementById('topicsList').innerHTML = '<div class="loading-spinner"></div>';
          document.getElementById('selectAllContainer').style.display = 'none';
          
          google.script.run
            .withSuccessHandler(showTopics)
            .withFailureHandler(showError)
            .getTopics(courseId);
        }

        function showTopics(topics) {
          const div = document.getElementById('topicsList');
          if (topics.length === 0) {
            div.innerHTML = '<p>No topics found in this course. All assignments will be downloaded.</p>';
            document.getElementById('downloadBtn').disabled = false;
            document.getElementById('selectAllContainer').style.display = 'none';
            document.getElementById('status').innerHTML = '';
            return;
          }

          div.innerHTML = topics.map(topic => \`
            <div class="topic-item">
              <label>
                <input type="checkbox" name="topic" value="\${topic.id}">
                \${topic.name}
              </label>
            </div>
          \`).join('');

          document.getElementById('downloadBtn').disabled = false;
          document.getElementById('status').innerHTML = '';
          document.getElementById('selectAllContainer').style.display = 'flex';
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
          const convertToPdf = document.getElementById('convertToPdf').checked;

          document.getElementById('status').innerHTML = '<div class="loading-spinner inline-spinner"></div> Preparing download...';
          document.getElementById('downloadBtn').disabled = true;
          document.getElementById('cancelBtn').style.display = 'block';
          document.getElementById('progress').style.display = 'block';
          document.getElementById('folderLink').style.display = 'none';
          document.getElementById('progressBar').style.width = '10%';
          downloadInProgress = true;
          
          // First get the folder URL
          google.script.run
            .withSuccessHandler(function(folderInfo) {
              folderUrl = folderInfo.url;
              document.getElementById('folderLink').innerHTML = 
                \`<a href="\${folderUrl}" target="_blank">Open folder in Google Drive</a>\`;
              document.getElementById('folderLink').style.display = 'block';
              document.getElementById('status').innerHTML = '<div class="loading-spinner inline-spinner"></div> Downloading assignments...';
              document.getElementById('progressBar').style.width = '30%';
              
              // Then start the actual download
              google.script.run
                .withSuccessHandler(showSuccess)
                .withFailureHandler(showError)
                .downloadAssignments(courseId, topicIds, folderInfo.id, convertToPdf);
            })
            .withFailureHandler(showError)
            .createDownloadFolder(courseId);
        }
        
        function cancelDownload() {
          if (!downloadInProgress) return;
          
          document.getElementById('status').innerHTML = '<div class="loading-spinner inline-spinner"></div> Cancelling download...';
          
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
          
          // Reset progress bar after 3 seconds
          setTimeout(() => {
            document.getElementById('progress').style.display = 'none';
            document.getElementById('progressBar').style.width = '0%';
          }, 3000);
        }

        function showError(error) {
          downloadInProgress = false;
          document.getElementById('status').innerHTML = 'Error: ' + error;
          document.getElementById('downloadBtn').disabled = false;
          document.getElementById('cancelBtn').style.display = 'none';
          document.getElementById('progress').style.display = 'none';
        }

        // In der script-Sektion, nach dem Laden der Seite
        google.script.run
          .withSuccessHandler(function(result) {
            console.log("Logged in as: " + (result.user || "unknown"));
          })
          .logAccess();
    </script>
    </body>
    </html>
  `);
  
  return html.setTitle('Classroom Assignment Downloader').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
function downloadAssignments(courseId, topicIds, rootFolderId, convertToPdf) {
  try {
    // Reset cancel flag
    shouldCancelDownload = false;
    
    // Get the root folder by ID
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    
    // Verify course exists
    const courseDetails = Classroom.Courses.get(courseId);
    const courseName = courseDetails.name;
    
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
      
      // Only add uncategorized folder if we have assignments without topics
      const hasUncategorizedAssignments = courseWork.some(work => 
        (!work.topicId || work.topicId === 'no-topic') && 
        (!topicIds.length || topicIds.includes(work.topicId))
      );
      
      if (hasUncategorizedAssignments) {
        topicFolders['no-topic'] = rootFolder.createFolder('Uncategorized');
      }
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
        
        // Get clean student name - ensure it's a string
        const rawStudentName = student.profile.name || student.profile.emailAddress || studentId.toString();
        const studentName = String(rawStudentName); // Force conversion to string
        
        // Log the raw student name and profile structure for debugging
        Logger.log("Raw student name: " + studentName);
        Logger.log("Student profile structure: " + JSON.stringify(student.profile));
        
        // Try to extract first and last name consistently
        let firstName = "";
        let lastName = "";
        
        // Check if we have structured name data
        if (student.profile.name && student.profile.name.givenName) {
          firstName = student.profile.name.givenName;
          lastName = student.profile.name.familyName || "";
        } 
        // Otherwise parse from the full name
        else {
          // Remove prefixes first
          const cleanedName = studentName.replace(/fullName|givenName|familyName/g, '');
          
          // Split by spaces and remove duplicates
          const nameParts = cleanedName.split(/\s+/).filter(part => part.trim().length > 0);
          const uniqueParts = [...new Set(nameParts)];
          
          if (uniqueParts.length >= 2) {
            firstName = uniqueParts[0];
            lastName = uniqueParts[uniqueParts.length - 1];
          } else if (uniqueParts.length === 1) {
            firstName = uniqueParts[0];
            lastName = "";
          }
        }
        
        // Construct a consistent name format: "FirstName LastName"
        let cleanStudentName = firstName;
        if (lastName) {
          cleanStudentName += " " + lastName;
        }
        
        // Clean up any remaining special characters
        cleanStudentName = cleanStudentName
          .replace(/[^\w\s-]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
        
        // If name is empty, use a fallback
        if (!cleanStudentName) {
          cleanStudentName = "Student-" + studentId;
        }
        
        // Log the final cleaned name
        Logger.log("Cleaned student name: " + cleanStudentName);
        
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
          
          // Include student name in filename
          const safeStudentNameForFile = cleanStudentName.replace(/\s+/g, '-');
          const baseFilename = `${safeAssignmentTitle}_${safeStudentNameForFile}`;
          
          if (attachment.driveFile) {
            downloadFile(attachment.driveFile, studentFolder, baseFilename, convertToPdf);
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

function downloadFile(driveFile, folder, baseFilename, convertToPdf) {
  try {
    const file = DriveApp.getFileById(driveFile.id);
    const fileName = file.getName();
    const mimeType = file.getMimeType();
    const originalExt = fileName.split('.').pop().toLowerCase();
    
    // Define convertible MIME types
    const convertibleMimeTypes = {
      'application/vnd.google-apps.document': 'application/pdf',
      'application/vnd.google-apps.spreadsheet': 'application/pdf',
      'application/vnd.google-apps.presentation': 'application/pdf',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'application/pdf',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'application/pdf',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'application/pdf',
      'application/msword': 'application/pdf',
      'application/vnd.ms-excel': 'application/pdf',
      'application/vnd.ms-powerpoint': 'application/pdf'
    };
    
    const shouldConvert = convertToPdf && convertibleMimeTypes.hasOwnProperty(mimeType);
    
    if (shouldConvert) {
      Logger.log(`Converting file ${fileName} (${mimeType}) to PDF`);
      try {
        // For Google Workspace files
        if (mimeType.includes('google-apps')) {
          const pdfBlob = file.getAs('application/pdf');
          const newFilename = baseFilename + '.pdf';
          folder.createFile(pdfBlob).setName(newFilename);
          Logger.log(`Successfully converted Google file to PDF: ${newFilename}`);
        } 
        // For Microsoft Office files
        else {
          const newFilename = baseFilename + '.pdf';
          const pdfFile = Drive.Files.copy(
            {title: newFilename, mimeType: 'application/pdf'},
            driveFile.id,
            {convert: true}
          );
          DriveApp.getFileById(pdfFile.id).moveTo(folder);
          Logger.log(`Successfully converted Office file to PDF: ${newFilename}`);
        }
      } catch (convError) {
        Logger.log(`Error converting to PDF: ${convError}. Falling back to original format.`);
        const newFilename = baseFilename + '.' + originalExt;
        file.makeCopy(newFilename, folder);
      }
    } else {
      // Keep original format
      const newFilename = baseFilename + '.' + originalExt;
      file.makeCopy(newFilename, folder);
      Logger.log(`Copied file in original format: ${newFilename}`);
    }
  } catch (error) {
    Logger.log('Error downloading file: ' + error);
  }
}

function createLinkFile(link, folder, baseFilename) {
  const content = 'URL: ' + link.url;
  const newFilename = baseFilename + '_link.txt';
  folder.createFile(newFilename, content);
}

/**
 * Creates the download folder and returns its URL
 */
function createDownloadFolder(courseId) {
  try {
    // Verify course exists
    const courseDetails = Classroom.Courses.get(courseId);
    const courseName = courseDetails.name;
    
    // Create main folder
    const rootFolder = DriveApp.createFolder('Classroom Downloads - ' + courseName);
    
    return {
      id: rootFolder.getId(),
      url: rootFolder.getUrl()
    };
  } catch (error) {
    Logger.log("Error creating folder: " + error);
    throw new Error("Failed to create download folder: " + error.message);
  }
}

/**
 * Logs user access to the app
 */
function logAccess() {
  try {
    const user = Session.getActiveUser().getEmail();
    Logger.log("App accessed by: " + user);
    return { success: true, user: user };
  } catch (error) {
    Logger.log("Error logging access: " + error);
    return { success: false, error: error.toString() };
  }
}