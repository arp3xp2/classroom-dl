function testApiStepByStep() {
  try {
    // Step 1: Test if Classroom API is available
    Logger.log("Step 1: Testing if Classroom API is available");
    if (typeof Classroom !== 'undefined') {
      Logger.log("✓ Classroom API is available");
    } else {
      Logger.log("✗ Classroom API is not available - make sure to add it in Services");
      return;
    }
    
    // Step 2: Try listing courses (without accessing .length)
    Logger.log("Step 2: Trying to list courses");
    try {
      const coursesResponse = Classroom.Courses.list();
      Logger.log("✓ Courses API call successful");
      Logger.log("Response type: " + typeof coursesResponse);
      
      if (coursesResponse) {
        if (coursesResponse.courses) {
          Logger.log("Number of courses: " + coursesResponse.courses.length);
          
          // Log the first course if available
          if (coursesResponse.courses.length > 0) {
            const firstCourse = coursesResponse.courses[0];
            Logger.log("First course ID: " + firstCourse.id);
            Logger.log("First course name: " + firstCourse.name);
            
            // Save this ID for further testing
            const testCourseId = firstCourse.id;
            
            // Step 3: Try getting a single course
            Logger.log("Step 3: Trying to get course details for ID: " + testCourseId);
            try {
              const courseDetails = Classroom.Courses.get(testCourseId);
              Logger.log("✓ Course details retrieved successfully");
              Logger.log("Course name: " + courseDetails.name);
              
              // Step 4: Try listing course work
              Logger.log("Step 4: Trying to list course work");
              try {
                const courseWorkResponse = Classroom.Courses.CourseWork.list(testCourseId);
                Logger.log("✓ Course work API call successful");
                
                if (courseWorkResponse && courseWorkResponse.courseWork) {
                  Logger.log("Number of course work items: " + courseWorkResponse.courseWork.length);
                } else {
                  Logger.log("No course work items found or unexpected response structure");
                }
              } catch (e) {
                Logger.log("✗ Error listing course work: " + e);
              }
              
              // Step 5: Try listing topics
              Logger.log("Step 5: Trying to access topics");
              try {
                Logger.log("Checking if Topics API endpoint exists");
                if (Classroom.Courses.Topics) {
                  Logger.log("Topics API endpoint found");
                  
                  const topicsResponse = Classroom.Courses.Topics.list(testCourseId);
                  Logger.log("Topics API call successful");
                  Logger.log("Response: " + JSON.stringify(topicsResponse));
                } else {
                  Logger.log("Topics API endpoint not found - might be a different structure");
                  
                  // Alternative: Check if topics are mentioned in course details
                  Logger.log("Checking if course details contains topic information");
                  Logger.log("Course details keys: " + Object.keys(courseDetails).join(", "));
                }
              } catch (e) {
                Logger.log("✗ Error accessing topics: " + e);
              }
            } catch (e) {
              Logger.log("✗ Error getting course details: " + e);
            }
          } else {
            Logger.log("No courses found to test with");
          }
        } else {
          Logger.log("Unexpected response structure - 'courses' property missing");
        }
      } else {
        Logger.log("Empty response from Courses.list()");
      }
    } catch (e) {
      Logger.log("✗ Error listing courses: " + e);
    }
  } catch (error) {
    Logger.log("Overall test error: " + error);
  }
}

function listTopicsForCourse(courseId) {
  try {
    // If courseId is a Base64 string, convert it
    const numericCourseId = courseId.length > 15 ? courseId : convertBase64ToID(courseId);
    
    Logger.log(`Attempting to list topics for course ID: ${numericCourseId}`);
    
    const response = Classroom.Courses.Topics.list(numericCourseId);
    
    if (response && response.topic) {
      Logger.log(`Topics for course ${numericCourseId}:`);
      for (const topic of response.topic) {
        Logger.log(`ID: ${topic.topicId}, Name: ${topic.name}`);
      }
      return response.topic;
    } else {
      Logger.log("No topics found in this course");
      return [];
    }
  } catch(error) {
    Logger.log("Error retrieving topics: " + error);
    Logger.log("Stack: " + error.stack);
    return [];
  }
}

function diagnoseCourseAccess(courseId) {
  try {
    // If courseId is a Base64 string, convert it
    const numericCourseId = courseId.length > 15 ? courseId : convertBase64ToID(courseId);
    
    Logger.log(`Diagnosing access for course ID: ${numericCourseId}`);
    
    // 1. Try to get the course details
    try {
      const courseDetails = Classroom.Courses.get(numericCourseId);
      Logger.log(`Course found: ${courseDetails.name}`);
    } catch (e) {
      Logger.log(`Error getting course details: ${e}`);
    }
    
    // 2. Try to list course work with minimal assumptions
    try {
      const courseWorkResponse = Classroom.Courses.CourseWork.list(numericCourseId);
      Logger.log(`Course work response type: ${typeof courseWorkResponse}`);
      Logger.log(`Course work response keys: ${courseWorkResponse ? Object.keys(courseWorkResponse).join(', ') : 'null'}`);
      
      if (courseWorkResponse && courseWorkResponse.courseWork) {
        Logger.log(`Found ${courseWorkResponse.courseWork.length} course work items`);
        
        // Log the first item to see its structure
        if (courseWorkResponse.courseWork.length > 0) {
          const firstItem = courseWorkResponse.courseWork[0];
          Logger.log(`First course work: ${JSON.stringify(firstItem)}`);
          Logger.log(`Topic ID: ${firstItem.topicId || 'None'}`);
        }
      } else {
        Logger.log('No course work found or unexpected response structure');
      }
    } catch (e) {
      Logger.log(`Error listing course work: ${e}`);
    }
    
    // 3. Try accessing topics directly
    try {
      // This API endpoint might not exist as expected
      const topicsResponse = Classroom.Courses.Topics.list(numericCourseId);
      Logger.log(`Topics response type: ${typeof topicsResponse}`);
      Logger.log(`Topics response: ${JSON.stringify(topicsResponse)}`);
    } catch (e) {
      Logger.log(`Error accessing topics API: ${e}`);
    }
    
  } catch (error) {
    Logger.log(`Overall diagnostic error: ${error}`);
  }
}

/**
 * Google Classroom Assignment Downloader
 * 
 * This script downloads all student assignments from a Google Classroom course,
 * organizes them by topic and student, and renames files using a consistent format.
 * 
 * Required scopes:
 * - https://www.googleapis.com/auth/classroom.courses.readonly
 * - https://www.googleapis.com/auth/classroom.coursework.students.readonly
 * - https://www.googleapis.com/auth/classroom.rosters.readonly
 * - https://www.googleapis.com/auth/drive
 */

/**
 * Convert Base64 URL ID to numeric ID
 */
function convertBase64ToID(base64Id) {
  // Convert Base64 URL-safe to standard Base64
  const standard = base64Id.replace(/-/g, '+').replace(/_/g, '/');
  
  // Decode Base64 to string (which will be the numeric ID)
  try {
    const decoded = Utilities.base64Decode(standard, Utilities.Charset.UTF_8);
    return Utilities.newBlob(decoded).getDataAsString();
  } catch (e) {
    Logger.log("Error decoding: " + e);
    return base64Id; // Return original if decoding fails
  }
}

/**
 * List all courses the user has access to
 */
function listAllCourses() {
  try {
    const courses = Classroom.Courses.list().courses;
    Logger.log("Verfügbare Kurse:");
    if (courses && courses.length > 0) {
      for (const course of courses) {
        Logger.log(`ID: ${course.id}, Name: ${course.name}`);
      }
    } else {
      Logger.log("Keine Kurse gefunden.");
    }
  } catch (error) {
    Logger.log("API-Fehler: " + error);
  }
}
function listTopicsFromCourseWork(courseId) {
  try {
    // If courseId is a Base64 string, convert it
    const numericCourseId = courseId.length > 15 ? courseId : convertBase64ToID(courseId);
    
    Logger.log(`Getting topics from course work for course ID: ${numericCourseId}`);
    
    // Get all course work
    const courseWork = getAllCourseWork(numericCourseId);
    
    // Extract unique topics from coursework
    const topicMap = {};
    for (const work of courseWork) {
      if (work.topicId) {
        topicMap[work.topicId] = work.topic || 'Unknown Topic Name';
      }
    }
    
    // Log the topics
    const topicIds = Object.keys(topicMap);
    if (topicIds.length > 0) {
      Logger.log(`Found ${topicIds.length} topics in course work:`);
      for (const topicId of topicIds) {
        Logger.log(`ID: ${topicId}, Name: ${topicMap[topicId]}`);
      }
    } else {
      Logger.log("No topics found in course work");
    }
    
    return topicMap;
  } catch(error) {
    Logger.log("Error getting topics from course work: " + error);
    Logger.log("Stack: " + error.stack);
    return {};
  }
}
/**
 * Run this function directly to download assignments with hardcoded values
 * This is useful if you can't deploy the web app
 */
function runManualDownload() {
  // Get the Base64 IDs from the URL
  const urlCourseId = 'Njc0Mjk5NjE4NTEz'; // From https://classroom.google.com/u/1/c/Njc0Mjk5NjE4NTEz
  const urlTopicId = 'NzAyNzU3MjI5NTYz'; // From /tc/NzAyNzU3MjI5NTYz
  
  // Convert to numeric IDs for the API
  const courseId = convertBase64ToID(urlCourseId);
  
  // Set up topics - use an empty object to download all topics
  const topics = {};
  
  // Uncomment to specify only certain topics
  topics[convertBase64ToID(urlTopicId)] = 'Admin / Zertifikate';
  
  // Download the assignments
  const result = downloadAssignments(courseId, topics);
  Logger.log(result);
}

/**
 * Creates a simple web UI for configuration
 */
function doGet() {
  const html = HtmlService.createHtmlOutput(`
    <h2>Google Classroom Assignment Downloader</h2>
    <p>Enter the Course ID from the URL (e.g., for https://classroom.google.com/u/1/c/Njc0Mjk5NjE4NTEz, enter "Njc0Mjk5NjE4NTEz")</p>
    <form id="configForm">
      <label>Course ID:</label><br>
      <input type="text" name="courseId" style="width:100%" required><br><br>
      
      <div id="topicsSection">
        <p>Topics (leave blank to download all)</p>
        <div id="topicsList">
          <div class="topicRow">
            <input type="text" name="topicId" placeholder="Topic ID" style="width:45%">
            <input type="text" name="topicName" placeholder="Folder Name" style="width:45%">
            <button type="button" onclick="addTopicRow()">+</button>
          </div>
        </div>
      </div>
      
      <br><br>
      <input type="submit" value="Download Assignments">
    </form>
    
    <script>
      function addTopicRow() {
        const div = document.createElement('div');
        div.className = 'topicRow';
        div.innerHTML = '<input type="text" name="topicId" placeholder="Topic ID" style="width:45%">' +
                       '<input type="text" name="topicName" placeholder="Folder Name" style="width:45%">' +
                       '<button type="button" onclick="this.parentNode.remove()">-</button>';
        document.getElementById('topicsList').appendChild(div);
      }
      
      document.getElementById('configForm').addEventListener('submit', function(e) {
        e.preventDefault();
        const formData = new FormData(this);
        const courseId = formData.get('courseId');
        
        const topics = {};
        const topicIds = formData.getAll('topicId');
        const topicNames = formData.getAll('topicName');
        
        for (let i = 0; i < topicIds.length; i++) {
          if (topicIds[i] && topicNames[i]) {
            topics[topicIds[i]] = topicNames[i];
          }
        }
        
        google.script.run
          .withSuccessHandler(function(result) {
            document.body.innerHTML = '<h2>Result</h2><p>' + result + '</p>';
          })
          .withFailureHandler(function(error) {
            document.body.innerHTML = '<h2>Error</h2><p>' + error + '</p>';
          })
          .downloadAssignments(courseId, topics);
      });
    </script>
  `);
  
  return html.setTitle('Classroom Downloader');
}

/**
 * Main function to download assignments
 */
function downloadAssignments(courseId, topics) {
  try {
    // Validate the courseId
    if (!courseId) {
      return 'Error: Please provide a valid Course ID';
    }
    
    // If the courseId is a Base64 string, convert it
    if (courseId.length < 15) {
      courseId = convertBase64ToID(courseId);
    }
    
    // Convert any Base64 topic IDs to numeric
    const convertedTopics = {};
    for (const [topicId, topicName] of Object.entries(topics)) {
      const convertedId = topicId.length < 15 ? convertBase64ToID(topicId) : topicId;
      convertedTopics[convertedId] = topicName;
    }
    
    // Create the main folder for this course
    const courseDetails = Classroom.Courses.get(courseId);
    const courseName = courseDetails.name;
    const rootFolder = DriveApp.createFolder('Classroom Downloads - ' + courseName);
    Logger.log('Created root folder: ' + rootFolder.getName());
    
    // Get all students in the course
    const students = getAllStudents(courseId);
    Logger.log('Found ' + students.length + ' students');
    
    // Get all course work
    const courseWork = getAllCourseWork(courseId);
    Logger.log('Found ' + courseWork.length + ' assignments');
    
    // Create a map of topic folders
    const topicFolders = createTopicFolders(rootFolder, convertedTopics, courseWork);
    
    // Process each assignment
    let downloadCount = 0;
    for (const assignment of courseWork) {
      const assignmentTitle = assignment.title;
      const topicId = assignment.topicId || 'no-topic';
      
      // Skip if we don't have this topic in our list (if topics were specified)
      if (Object.keys(convertedTopics).length > 0 && !convertedTopics[topicId]) {
        continue;
      }
      
      const topicFolder = topicFolders[topicId];
      if (!topicFolder) {
        Logger.log('Skipping assignment with unknown topic: ' + assignmentTitle);
        continue;
      }
      
      // Get student submissions for this assignment
      const submissions = getAllSubmissions(courseId, assignment.id);
      Logger.log('Processing ' + submissions.length + ' submissions for: ' + assignmentTitle);
      
      // Process each student submission
      for (const submission of submissions) {
        const studentId = submission.userId;
        const student = students.find(s => s.userId === studentId);
        
        if (!student) {
          Logger.log('Unknown student ID: ' + studentId);
          continue;
        }
        
        // Only process submissions with attachments
        if (!submission.assignmentSubmission || !submission.assignmentSubmission.attachments) {
          continue;
        }
        
        // Create student folder if it doesn't exist
        let studentFolder = null;
        const studentFolders = topicFolder.getFoldersByName(student.profile.name);
        if (studentFolders.hasNext()) {
          studentFolder = studentFolders.next();
        } else {
          studentFolder = topicFolder.createFolder(student.profile.name);
        }
        
        // Process attachments
        const attachments = submission.assignmentSubmission.attachments;
        for (let i = 0; i < attachments.length; i++) {
          const attachment = attachments[i];
          
          // Format the student name for file naming
          const formattedStudentName = student.profile.name.toLowerCase().replace(/\s+/g, '-');
          
          // Create base filename
          let baseFilename = assignmentTitle + '_' + formattedStudentName;
          
          // Add number if multiple files
          if (attachments.length > 1) {
            baseFilename += '_' + (i + 1);
          }
          
          // Handle different attachment types
          if (attachment.driveFile) {
            downloadDriveFile(attachment.driveFile, studentFolder, baseFilename);
            downloadCount++;
          } else if (attachment.link) {
            createLinkFile(attachment.link, studentFolder, baseFilename);
            downloadCount++;
          } else if (attachment.youTubeVideo) {
            createYouTubeLink(attachment.youTubeVideo, studentFolder, baseFilename);
            downloadCount++;
          } else if (attachment.form) {
            createFormLink(attachment.form, studentFolder, baseFilename);
            downloadCount++;
          }
        }
      }
    }
    
    return `Downloaded ${downloadCount} files to ${rootFolder.getName()}. URL: ${rootFolder.getUrl()}`;
  } catch (error) {
    Logger.log('Error: ' + error);
    return 'Error: ' + error.toString();
  }
}

/**
 * Get all students in a course
 */
function getAllStudents(courseId) {
  const students = [];
  let pageToken = null;
  
  do {
    const response = Classroom.Courses.Students.list(courseId, { pageToken: pageToken });
    if (response.students) {
      students.push(...response.students);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return students;
}

/**
 * Get all coursework in a course
 */
function getAllCourseWork(courseId) {
  const courseWork = [];
  let pageToken = null;
  
  do {
    const response = Classroom.Courses.CourseWork.list(courseId, { pageToken: pageToken });
    if (response.courseWork) {
      courseWork.push(...response.courseWork);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return courseWork;
}

/**
 * Get all submissions for a coursework
 */
function getAllSubmissions(courseId, courseWorkId) {
  const submissions = [];
  let pageToken = null;
  
  do {
    const response = Classroom.Courses.CourseWork.StudentSubmissions.list(
      courseId, 
      courseWorkId, 
      { pageToken: pageToken }
    );
    if (response.studentSubmissions) {
      submissions.push(...response.studentSubmissions);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return submissions;
}

/**
 * Create folders for each topic
 */
function createTopicFolders(rootFolder, configTopics, courseWork) {
  const topicFolders = {};
  
  // If no topics specified, get all topics from the course
  if (Object.keys(configTopics).length === 0) {
    // Extract unique topics from coursework
    const uniqueTopics = [...new Set(courseWork.map(cw => cw.topicId).filter(Boolean))];
    
    // Create a default mapping
    for (const topicId of uniqueTopics) {
      configTopics[topicId] = 'Topic ' + topicId;
    }
    
    // Add a catch-all for assignments without topics
    configTopics['no-topic'] = 'Uncategorized';
  } else {
    // Add the no-topic category if not explicitly included
    if (!configTopics['no-topic']) {
      configTopics['no-topic'] = 'Uncategorized';
    }
  }
  
  // Create a folder for each topic
  for (const [topicId, topicName] of Object.entries(configTopics)) {
    const folder = rootFolder.createFolder(topicName);
    topicFolders[topicId] = folder;
    Logger.log('Created topic folder: ' + topicName);
  }
  
  return topicFolders;
}

/**
 * Download a Drive file and save it to the student folder
 */
function downloadDriveFile(driveFile, studentFolder, baseFilename) {
  try {
    const fileId = driveFile.id;
    const file = DriveApp.getFileById(fileId);
    
    // Get file extension
    let fileName = file.getName();
    const ext = fileName.split('.').pop();
    
    // Create new filename with extension
    const newFilename = baseFilename + '.' + ext;
    
    // Copy file to student folder
    file.makeCopy(newFilename, studentFolder);
    Logger.log('Downloaded: ' + newFilename);
  } catch (error) {
    Logger.log('Error downloading file: ' + error);
  }
}

/**
 * Create a text file with a link
 */
function createLinkFile(link, studentFolder, baseFilename) {
  const content = 'URL: ' + link.url;
  const newFilename = baseFilename + '_link.txt';
  studentFolder.createFile(newFilename, content);
  Logger.log('Created link file: ' + newFilename);
}

/**
 * Create a text file with a YouTube link
 */
function createYouTubeLink(youTubeVideo, studentFolder, baseFilename) {
  const content = 'YouTube Video: https://www.youtube.com/watch?v=' + youTubeVideo.id;
  const newFilename = baseFilename + '_youtube.txt';
  studentFolder.createFile(newFilename, content);
  Logger.log('Created YouTube link file: ' + newFilename);
}

/**
 * Create a text file with a Form link
 */
function createFormLink(form, studentFolder, baseFilename) {
  const content = 'Google Form: ' + form.formUrl;
  const newFilename = baseFilename + '_form.txt';
  studentFolder.createFile(newFilename, content);
  Logger.log('Created Form link file: ' + newFilename);
}