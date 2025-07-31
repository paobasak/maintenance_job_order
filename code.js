/**
 * Google Apps Script for Venue Reservation Form - CORS Free, FormData Style
 */
const RESERVATION_SHEET_NAME = 'Reservation';
const LOGS_SHEET_NAME = 'Logs';
const VENUES_SHEET_NAME = 'Venues';
const USERS_SHEET_NAME = 'Users';
const MESSAGES_SHEET_NAME = 'Messages';
const DEPARTMENTS_SHEET_NAME = 'Departments';
const AC_CONCERNS_SHEET_NAME = 'AC-Concerns';
const ANNOUNCEMENTS_SHEET_NAME = 'Announcements';

function doPost(e) {
  try {
    // Accept FormData (multipart/form-data) and get fields from e.parameter
    const data = e.parameter || {};
    const action = data.action || 'addReservation';

    // Handle user authentication
    if (action === 'login') {
      // Validate required fields
      if (!data.username || !data.password) {
        return createCorsResponse({ success: false, error: 'Missing username or password' });
      }

      const result = authenticateUser(data.username, data.password);
      return createCorsResponse(result);
    }

    // Handle user registration
    if (action === 'signup') {
      // Validate required fields for signup
      if (!data.username || !data.password || !data.fullName || !data.department || !data.userType || !data.email) {
        return createCorsResponse({ success: false, error: 'Missing required fields: username, password, fullName, department, userType, or email' });
      }

      const result = registerUser(data);
      return createCorsResponse(result);
    }

    // Handle profile update
    if (action === 'updateUserProfile') {
      // Validate required fields for profile update
      if (!data.username) {
        return createCorsResponse({ success: false, error: 'Username is required for profile update' });
      }

      const result = updateUserProfile(data);
      return createCorsResponse(result);
    }

    // Handle get user profile
    if (action === 'getUserProfile') {
      if (!data.username) {
        return createCorsResponse({ success: false, error: 'Username is required' });
      }

      const result = getUserProfile(data.username);
      return createCorsResponse(result);
    }

    // Handle send message
    if (action === 'sendMessage') {
      if (!data.senderUsername || !data.receiverUsername || !data.message) {
        return createCorsResponse({ success: false, error: 'Missing required fields: senderUsername, receiverUsername, or message' });
      }

      const result = sendMessage(data);
      return createCorsResponse(result);
    }

    // Handle get messages
    if (action === 'getMessages') {
      if (!data.user1 || !data.user2) {
        return createCorsResponse({ success: false, error: 'Missing required fields: user1 or user2' });
      }

      const result = getMessages(data.user1, data.user2);
      return createCorsResponse(result);
    }

    // Handle get all users
    if (action === 'getAllUsers') {
      const result = getAllUsers();
      return createCorsResponse(result);
    }

    // Handle mark messages as read
    if (action === 'markMessagesAsRead') {
      if (!data.currentUser || !data.otherUser) {
        return createCorsResponse({ success: false, error: 'Missing required fields: currentUser or otherUser' });
      }

      const result = markMessagesAsRead(data.currentUser, data.otherUser);
      return createCorsResponse(result);
    }

    // Handle get unread message counts
    if (action === 'getUnreadCounts') {
      if (!data.currentUser) {
        return createCorsResponse({ success: false, error: 'Missing required field: currentUser' });
      }

      const result = getUnreadMessageCounts(data.currentUser);
      return createCorsResponse(result);
    }

    // Handle get users with messages (batch endpoint)
    if (action === 'getUsersWithMessages') {
      const currentUser = data.currentUser || '';
      
      // Get all users
      const usersResult = getAllUsers();
      if (!usersResult.success) {
        return createCorsResponse(usersResult);
      }
      
      // Get unread counts if currentUser is provided
      let unreadCounts = {};
      if (currentUser) {
        const unreadResult = getUnreadMessageCounts(currentUser);
        if (unreadResult.success) {
          unreadCounts = unreadResult.unreadCounts;
        }
      }
      
      return createCorsResponse({
        success: true,
        users: usersResult.users,
        unreadCounts: unreadCounts
      });
    }

    // Handle add usage log
    if (action === 'addUsageLog') {
      if (!data.venue || !data.activityDate || !data.department) {
        return createCorsResponse({ success: false, error: 'Missing required fields: venue, activityDate, or department' });
      }

      const result = addUsageLog(data);
      return createCorsResponse(result);
    }

    // Handle update usage log
    if (action === 'updateUsageLog') {
      if (!data.logId || !data.venue || !data.activityDate || !data.department) {
        return createCorsResponse({ success: false, error: 'Missing required fields: logId, venue, activityDate, or department' });
      }

      const result = updateUsageLog(data);
      return createCorsResponse(result);
    }

    // Handle get usage logs
    if (action === 'getUsageLogs') {
      const result = getUsageLogs();
      return createCorsResponse(result);
    }

    // Handle get all reservations for calendar
    if (action === 'getAllReservations') {
      const result = getReservations('getAllReservations');
      return createCorsResponse(result);
    }

    // Handle get departments
    if (action === 'getDepartments') {
      const result = getDepartments();
      return createCorsResponse(result);
    }

    // Handle get venues
    if (action === 'getVenues') {
      const result = getVenues();
      return createCorsResponse(result);
    }

    // Handle manage users (admin only)
    if (action === 'getUsers') {
      const result = getUsers();
      return createCorsResponse(result);
    }

    // Handle update user status (admin only)
    if (action === 'updateUserStatus') {
      if (!data.username || !data.status) {
        return createCorsResponse({ success: false, error: 'Missing required fields: username or status' });
      }

      const result = updateUserStatus(data);
      return createCorsResponse(result);
    }

    // Handle delete user (admin only)
    if (action === 'deleteUser') {
      if (!data.username) {
        return createCorsResponse({ success: false, error: 'Missing required field: username' });
      }

      const result = deleteUser(data.username);
      return createCorsResponse(result);
    }

    // Handle update user venue assignments (admin only)
    if (action === 'updateUserVenueAssignments') {
      if (!data.username) {
        return createCorsResponse({ success: false, error: 'Missing required field: username' });
      }

      const result = updateUserVenueAssignments(data);
      return createCorsResponse(result);
    }

    // Handle job order submission
    if (action === 'submitJobOrder') {
      // Validate required fields (no form number needed, it will be generated)
      if (!data.department || !data.location || !data.natureOfWork || !data.requestedBy) {
        return createCorsResponse({ success: false, error: 'Missing required fields: department, location, natureOfWork, or requestedBy' });
      }

      const result = submitJobOrder(data);
      return createCorsResponse(result);
    }

    // Handle get job orders
    if (action === 'getJobOrders') {
      const result = getJobOrders(data.username);
      return createCorsResponse(result);
    }

    // Handle update job order
    if (action === 'updateJobOrder') {
      Logger.log('DEBUG: updateJobOrder action detected');
      // Validate required fields
      if (!data.formNumber || !data.location || !data.natureOfWork) {
        return createCorsResponse({ success: false, error: 'Missing required fields: formNumber, location, or natureOfWork' });
      }

      Logger.log(`DEBUG: Calling updateJobOrder with formNumber=${data.formNumber}, location=${data.location}, natureOfWork=${data.natureOfWork}`);
      const result = updateJobOrder(data.formNumber, data.location, data.natureOfWork);
      Logger.log(`DEBUG: updateJobOrder result: ${JSON.stringify(result)}`);
      return createCorsResponse(result);
    }

    // Handle aircon concern submission
    if (action === 'submitAirconConcern') {
      // Validate required fields
      if (!data.date || !data.department || !data.airconType || !data.concerns || !data.reportedBy) {
        return createCorsResponse({ success: false, error: 'Missing required fields: date, department, airconType, concerns, or reportedBy' });
      }

      const result = submitAirconConcern(data);
      return createCorsResponse(result);
    }

    // Handle get AC concerns
    if (action === 'getACConcerns') {
      const result = getACConcerns(data.department);
      return createCorsResponse(result);
    }

    // Handle update AC concern
    if (action === 'updateACConcern') {
      // Validate required fields
      if (!data.concernId || !data.status) {
        return createCorsResponse({ success: false, error: 'Missing required fields: concernId or status' });
      }

      const result = updateACConcern(data.concernId, data.concerns, data.status, data.dateAccomplished);
      return createCorsResponse(result);
    }

    // Handle get announcement
    if (action === 'getAnnouncement') {
      const result = getAnnouncement();
      return createCorsResponse(result);
    }

    // Handle reservation submission
    if (action === 'addReservation') {
      // Validate required fields
      if (!data.venue || !data.activityDate || !data.department || !data.purpose) {
        return createCorsResponse({ success: false, error: 'Missing required fields: venue, activityDate, department, or purpose' });
      }

      // Add the reservation to Google Sheets
      const result = addReservationToSheet(data);

      // You could add email notification here if needed

      return createCorsResponse(result);
    }

    // Handle venue addition
    if (action === 'addVenue') {
      // Validate required fields for venue
      if (!data.venueName || !data.capacity) {
        return createCorsResponse({ success: false, error: 'Missing required fields: venueName or capacity' });
      }

      const result = addVenueToSheet(data);
      return createCorsResponse(result);
    }

    return createCorsResponse({ success: false, error: 'Unknown action' });
  } catch (error) {
    return createCorsResponse({ success: false, error: error.toString() });
  }
}

// GET (optional, for viewing reservations - can be expanded)
function doGet(e) {
  try {
    const action = e.parameter.action || '';
    
    if (action === 'getVenues') {
      return createCorsResponse(getVenues());
    }
    
    if (action === 'getDepartments') {
      return createCorsResponse(getDepartments());
    }
    
    if (action === 'getPending' || action === 'getApproved' || action === 'getAllReservations') {
      return createCorsResponse(getReservations(action));
    }
    
    return createCorsResponse({ success: true, message: 'Venue Reservation API is running.' });
  } catch (error) {
    return createCorsResponse({ success: false, error: error.toString() });
  }
}

// Preflight for CORS
function doOptions(e) {
  return createCorsResponse({});
}

// Helper for CORS JSON response
function createCorsResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// Add reservation to Google Sheet
function addReservationToSheet(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(RESERVATION_SHEET_NAME);
      sheet.appendRow([
        'Timestamp','Reservation ID','Venue','Activity Date','Day','Department','Time','Date Submitted','Purpose','Facilities','Status'
      ]);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 11);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 150); // Timestamp
      sheet.setColumnWidth(2, 120); // Reservation ID
      sheet.setColumnWidth(3, 150); // Venue
      sheet.setColumnWidth(4, 120); // Activity Date
      sheet.setColumnWidth(5, 100); // Day
      sheet.setColumnWidth(6, 150); // Department
      sheet.setColumnWidth(7, 120); // Time
      sheet.setColumnWidth(8, 150); // Date Submitted
      sheet.setColumnWidth(9, 200); // Purpose
      sheet.setColumnWidth(10, 200); // Facilities
      sheet.setColumnWidth(11, 100); // Status
    }
    const reservationId = 'VR-' + Date.now();
    const timestamp = new Date();
    
    // Parse and format the activity date properly
    let activityDate = data.activityDate;
    if (activityDate) {
      // If it's a date string in YYYY-MM-DD format, keep it as string to avoid timezone issues
      if (typeof activityDate === 'string' && activityDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
        // Keep as string - don't convert to Date object to avoid timezone problems
        console.log('Keeping reservation date as string to avoid timezone issues:', activityDate);
      } else if (typeof activityDate === 'string') {
        // Handle other date formats by converting to Date and then back to YYYY-MM-DD string
        const parsedDate = new Date(activityDate);
        if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() >= 1900 && parsedDate.getFullYear() <= 2100) {
          // Convert back to YYYY-MM-DD string to maintain consistency
          const year = parsedDate.getFullYear();
          const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
          const day = String(parsedDate.getDate()).padStart(2, '0');
          activityDate = `${year}-${month}-${day}`;
          console.log('Converted reservation date to YYYY-MM-DD string:', activityDate);
        } else {
          console.log('Could not parse reservation date string:', activityDate);
          activityDate = null;
        }
      } else if (activityDate instanceof Date) {
        // If it's a Date object, convert to YYYY-MM-DD string
        if (!isNaN(activityDate.getTime()) && activityDate.getFullYear() >= 1900 && activityDate.getFullYear() <= 2100) {
          const year = activityDate.getFullYear();
          const month = String(activityDate.getMonth() + 1).padStart(2, '0');
          const day = String(activityDate.getDate()).padStart(2, '0');
          activityDate = `${year}-${month}-${day}`;
          console.log('Converted reservation Date object to YYYY-MM-DD string:', activityDate);
        } else {
          console.log('Invalid reservation Date object:', activityDate);
          activityDate = null;
        }
      }
    }
    
    // Parse dateSubmitted if provided, otherwise use current timestamp
    let dateSubmitted = data.dateSubmitted ? new Date(data.dateSubmitted) : timestamp;
    if (isNaN(dateSubmitted.getTime())) {
      dateSubmitted = timestamp;
    }
    
    sheet.appendRow([
      timestamp,
      reservationId,
      data.venue,
      activityDate,
      data.day || '',
      data.department,
      data.time || '',
      dateSubmitted,
      data.purpose,
      data.facilities || '',
      'Approved'
    ]);
    return {
      success: true,
      message: 'Reservation added successfully',
      reservationId: reservationId,
      timestamp: timestamp.toISOString()
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Get venues from Venues sheet
function getVenues() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(VENUES_SHEET_NAME);
    
    if (!sheet) {
      // Create Venues sheet with default venues
      sheet = ss.insertSheet(VENUES_SHEET_NAME);
      sheet.appendRow(['Venue Name', 'Capacity', 'Facilities', 'Status']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 200); // Venue Name
      sheet.setColumnWidth(2, 100); // Capacity
      sheet.setColumnWidth(3, 300); // Facilities
      sheet.setColumnWidth(4, 100); // Status
      
      // Add default venues
      sheet.appendRow(['Conference Room A', '20', 'Projector, Whiteboard, Audio System', 'Active']);
      sheet.appendRow(['Conference Room B', '15', 'TV Screen, Whiteboard', 'Active']);
      sheet.appendRow(['Main Hall', '100', 'Stage, Sound System, Projector', 'Active']);
      sheet.appendRow(['Training Room 1', '30', 'Computers, Projector', 'Active']);
      sheet.appendRow(['Training Room 2', '25', 'Projector, Whiteboard', 'Active']);
      sheet.appendRow(['Boardroom', '12', 'Conference Table, TV Screen', 'Active']);
    }
    
    const data = sheet.getDataRange().getValues();
    const venues = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[3] === 'Active') { // Venue Name exists and Status is Active
        venues.push({
          name: row[0],
          capacity: row[1] || '',
          facilities: row[2] || '',
          status: row[3] || 'Active'
        });
      }
    }
    
    return {
      success: true,
      venues: venues
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Get departments from Departments sheet
function getDepartments() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DEPARTMENTS_SHEET_NAME);
    
    if (!sheet) {
      // Create Departments sheet with default departments
      sheet = ss.insertSheet(DEPARTMENTS_SHEET_NAME);
      sheet.appendRow(['Department ID', 'Department Name']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 2);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 100); // Department ID
      sheet.setColumnWidth(2, 300); // Department Name
      
      // Add all the departments with incremental IDs
      const departments = [
        'ACCOUNTING BASAK',
        'ADMIN',
        'ATHLETICS',
        'BOOKSTORE',
        'CMO BASAK',
        'COLLEGE LIBRARY',
        'COLLEGE OF EDUCATION',
        'DIRECTOR\'S OFFICE',
        'DRIVER',
        'ELEMENTARY LIBRARY',
        'ELEMENTARY OFFICE',
        'HEALTH SERVICES DEPT. DENTAL',
        'HEALTH SERVICES DEPT. MEDICAL',
        'HRMO/ IMS',
        'IMC - BASAK',
        'JHS LIBRARY',
        'JHS MAPEH',
        'JHS OFFICE',
        'JHS SCI. LAB.',
        'JHS TLE - DRAFTING',
        'JHS TLE – SERVER',
        'MAINTENANCE',
        'OSAS',
        'PACUBAS',
        'PAIR OFFICE',
        'PAO BASAK',
        'PE COLLEGE',
        'PRINTING SECTION',
        'RECORDS BASAK',
        'RIDEM',
        'RITTC',
        'ROTC',
        'SCS - CICCT',
        'SDPC LA SALUD',
        'SDPC OLCON',
        'SDPC SEM',
        'SHS LIBRARY',
        'SHS OFFICE',
        'SPED CENTER',
        'SSC',
        'SSD',
        'UITSS / ITCNDC'
      ];
      
      // Add departments with IDs starting from 1
      departments.forEach((dept, index) => {
        sheet.appendRow([index + 1, dept]);
      });
    }
    
    const data = sheet.getDataRange().getValues();
    const departments = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[1]) { // Department ID and Name
        departments.push({
          id: row[0],
          name: row[1]
        });
      }
    }
    
    return {
      success: true,
      departments: departments
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Get reservations (placeholder for future expansion)
function getReservations(action) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    
    if (!sheet) {
      return { success: true, reservations: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const reservations = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Has timestamp
        // Handle date values properly
        let timestamp = row[0];
        let activityDate = row[3];
        let dateSubmitted = row[7];
        
        // Convert timestamps to proper Date objects if they're not already
        if (timestamp && !(timestamp instanceof Date)) {
          timestamp = new Date(timestamp);
        }
        
        // Handle activity date - keep as string if it's in YYYY-MM-DD format
        if (activityDate) {
          if (typeof activityDate === 'string' && activityDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
            // Already in correct YYYY-MM-DD format, keep as string
            console.log('Reservation activity date is already in YYYY-MM-DD format:', activityDate);
          } else if (activityDate instanceof Date) {
            // Convert Date object to YYYY-MM-DD string
            if (activityDate.getFullYear() >= 1900 && activityDate.getFullYear() <= 2100) {
              const year = activityDate.getFullYear();
              const month = String(activityDate.getMonth() + 1).padStart(2, '0');
              const day = String(activityDate.getDate()).padStart(2, '0');
              activityDate = `${year}-${month}-${day}`;
              console.log('Converted reservation activity date from Date to YYYY-MM-DD:', activityDate);
            } else {
              console.log('Invalid reservation activity date year:', activityDate.getFullYear(), 'Setting to null');
              activityDate = null;
            }
          } else if (typeof activityDate === 'number') {
            // Handle Google Sheets serial number (if any)
            if (activityDate > 25569) { // Roughly 1970-01-01 in Excel serial date
              const dateObj = new Date((activityDate - 25569) * 86400 * 1000);
              if (dateObj.getFullYear() >= 1900 && dateObj.getFullYear() <= 2100) {
                const year = dateObj.getFullYear();
                const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                const day = String(dateObj.getDate()).padStart(2, '0');
                activityDate = `${year}-${month}-${day}`;
                console.log('Converted reservation serial number to YYYY-MM-DD:', activityDate);
              } else {
                activityDate = null;
              }
            } else {
              activityDate = null;
            }
          } else {
            // Try to parse as string and convert to YYYY-MM-DD
            const dateObj = new Date(activityDate);
            if (!isNaN(dateObj.getTime()) && dateObj.getFullYear() >= 1900 && dateObj.getFullYear() <= 2100) {
              const year = dateObj.getFullYear();
              const month = String(dateObj.getMonth() + 1).padStart(2, '0');
              const day = String(dateObj.getDate()).padStart(2, '0');
              activityDate = `${year}-${month}-${day}`;
              console.log('Parsed and converted reservation to YYYY-MM-DD:', activityDate);
            } else {
              console.log('Could not parse reservation activity date:', activityDate, 'Setting to null');
              activityDate = null;
            }
          }
        }
        
        if (dateSubmitted && !(dateSubmitted instanceof Date)) {
          dateSubmitted = new Date(dateSubmitted);
          // Check for problematic dates
          if (isNaN(dateSubmitted.getTime()) || dateSubmitted.getFullYear() < 1900) {
            dateSubmitted = null;
          }
        }
        
        const reservation = {
          timestamp: timestamp ? timestamp.toISOString() : null,
          reservationId: row[1],
          venue: row[2],
          activityDate: activityDate, // Already in YYYY-MM-DD format as string
          day: row[4],
          department: row[5],
          time: row[6],
          dateSubmitted: dateSubmitted ? dateSubmitted.toISOString() : null,
          purpose: row[8],
          facilities: row[9],
          status: row[10]
        };
        
        // Filter by status if needed
        if (action === 'getPending' && reservation.status === 'Pending') {
          reservations.push(reservation);
        } else if (action === 'getApproved' && reservation.status === 'Approved') {
          reservations.push(reservation);
        } else if (action === 'getAllReservations' || !action || action === 'getAll') {
          reservations.push(reservation);
        }
      }
    }
    
    return {
      success: true,
      reservations: reservations
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Add venue to Venues sheet
function addVenueToSheet(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(VENUES_SHEET_NAME);
    
    if (!sheet) {
      // Create Venues sheet with headers if it doesn't exist
      sheet = ss.insertSheet(VENUES_SHEET_NAME);
      sheet.appendRow(['Venue Name', 'Capacity', 'Facilities', 'Status']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 200); // Venue Name
      sheet.setColumnWidth(2, 100); // Capacity
      sheet.setColumnWidth(3, 300); // Facilities
      sheet.setColumnWidth(4, 100); // Status
    }
    
    // Check if venue name already exists
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] && existingData[i][0].toString().toLowerCase() === data.venueName.toLowerCase()) {
        return { success: false, error: 'Venue name already exists' };
      }
    }
    
    // Add new venue
    sheet.appendRow([
      data.venueName,
      data.capacity,
      data.facilities || '',
      'Active'
    ]);
    
    return {
      success: true,
      message: 'Venue added successfully'
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Authenticate user
function authenticateUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      // Create Users sheet with default admin user
      sheet = createUsersSheet(ss);
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[1]) { // Username and password exist
        if (row[0].toString().toLowerCase() === username.toLowerCase() && 
            row[1].toString() === password && 
            row[7] === 'Active') { // Check if user is active (Status is now column 7)
          
          const userType = row[6]; // Get actual user type from sheet without default fallback
          // Only 'admin' and 'user' types should have venue assignments
          const venueAssignments = (userType === 'admin' || userType === 'user') ? (row[10] || 'ALL') : '';
          
          return {
            success: true,
            message: 'Login successful',
            user: {
              username: row[0] || '',
              fullName: row[2] || '',
              department: row[3] || '',
              email: row[4] || '',
              phone: (row[5] && row[5].toString() !== '#ERROR!') ? row[5].toString() : '',
              userType: userType,
              status: row[7] || 'Active',
              profilePicture: row[8] || '',
              venueAssignments: venueAssignments
            }
          };
        }
      }
    }
    
    return { success: false, error: 'Invalid username or password' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Register new user
function registerUser(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      // Create Users sheet with default admin user
      sheet = createUsersSheet(ss);
    }
    
    // Check if username already exists
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] && existingData[i][0].toString().toLowerCase() === data.username.toLowerCase()) {
        return { success: false, error: 'Username already exists' };
      }
      if (existingData[i][4] && existingData[i][4].toString().toLowerCase() === data.email.toLowerCase()) {
        return { success: false, error: 'Email already exists' };
      }
    }
    
    // Add new user with the provided user type from form
    const timestamp = new Date();
    const userType = data.userType || 'student'; // Use provided user type, default to 'student' if not provided
    
    // Determine venue assignments based on user type
    // Only 'admin' and 'user' types get venue assignments, others get empty string
    const venueAssignments = (userType === 'admin' || userType === 'user') ? 'ALL' : '';
    
    sheet.appendRow([
      data.username,
      data.password,
      data.fullName,
      data.department,
      data.email,
      data.phone || '', // Phone number
      userType, // User type
      'Active',
      '', // No profile picture initially
      timestamp,
      venueAssignments // Venue assignments based on user type
    ]);
    
    return {
      success: true,
      message: 'User registered successfully',
      user: {
        username: data.username,
        fullName: data.fullName,
        department: data.department,
        email: data.email,
        phone: data.phone || '',
        userType: 'user',
        status: 'Active',
        profilePicture: ''
      }
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Create Users sheet with headers and default admin user
function createUsersSheet(ss) {
  const sheet = ss.insertSheet(USERS_SHEET_NAME);
  
  // Add headers - added Department column before Email
  sheet.appendRow(['Username', 'Password', 'Full Name', 'Department', 'Email', 'Phone', 'User Type', 'Status', 'Profile Picture', 'Created Date', 'Venue Assignments']);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, 11);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f4f6');
  headerRange.setBorder(true, true, true, true, true, true);
  
  // Set column widths
  sheet.setColumnWidth(1, 120); // Username
  sheet.setColumnWidth(2, 120); // Password
  sheet.setColumnWidth(3, 180); // Full Name
  sheet.setColumnWidth(4, 150); // Department
  sheet.setColumnWidth(5, 200); // Email
  sheet.setColumnWidth(6, 120); // Phone
  sheet.setColumnWidth(7, 100); // User Type
  sheet.setColumnWidth(8, 100); // Status
  sheet.setColumnWidth(9, 150); // Profile Picture
  sheet.setColumnWidth(10, 150); // Created Date
  sheet.setColumnWidth(11, 250); // Venue Assignments
  
  // Add default admin user
  const timestamp = new Date();
  sheet.appendRow([
    'admin',
    'admin',
    'System Administrator',
    'PAO BASAK', // Use a department from the actual departments list
    'admin@reservehub.com',
    '+1 (555) 123-4567',
    'admin',
    'Active',
    '', // No profile picture initially
    timestamp,
    'ALL' // Admin can see all venues
  ]);
  
  return sheet;
}

// Get user profile data
function getUserProfile(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[0].toString().toLowerCase() === username.toLowerCase()) {
        const userType = row[6]; // Get actual user type from sheet without default fallback
        // Only 'admin' and 'user' types should have venue assignments
        const venueAssignments = (userType === 'admin' || userType === 'user') ? (row[10] || 'ALL') : '';
        
        return {
          success: true,
          user: {
            username: row[0] || '',
            fullName: row[2] || '',
            department: row[3] || '',
            email: row[4] || '',
            phone: (row[5] && row[5].toString() !== '#ERROR!') ? row[5].toString() : '',
            userType: userType,
            status: row[7] || 'Active',
            profilePicture: row[8] || '',
            venueAssignments: venueAssignments
          }
        };
      }
    }
    
    return { success: false, error: 'User not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Update user profile
function updateUserProfile(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    const sheetData = sheet.getDataRange().getValues();
    
    // Find the user row
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      if (row[0] && row[0].toString().toLowerCase() === data.username.toLowerCase()) {
        // Update user data with proper validation
        if (data.fullName) {
          sheet.getRange(i + 1, 3).setValue(data.fullName);
        }
        if (data.department) {
          sheet.getRange(i + 1, 4).setValue(data.department);
        }
        if (data.email) {
          sheet.getRange(i + 1, 5).setValue(data.email);
        }
        // Handle phone field properly - allow empty string
        if (data.phone !== undefined) {
          sheet.getRange(i + 1, 6).setValue(data.phone || '');
        }
        // Handle userType field properly (admin only can change this)
        if (data.userType !== undefined) {
          sheet.getRange(i + 1, 7).setValue(data.userType);
          
          // Update venue assignments based on user type
          // Only 'admin' and 'user' types get venue assignments, others get empty string
          const venueAssignments = (data.userType === 'admin' || data.userType === 'user') ? 'ALL' : '';
          sheet.getRange(i + 1, 11).setValue(venueAssignments); // Column 11 is venue assignments
        }
        // Handle profile picture properly
        if (data.profilePicture !== undefined) {
          sheet.getRange(i + 1, 9).setValue(data.profilePicture || '');
        }
        
        // Get updated user data (need to include venue assignments)
        const updatedRow = sheet.getRange(i + 1, 1, 1, 11).getValues()[0];
        const userType = updatedRow[6]; // Get actual user type from sheet without default fallback
        // Only 'admin' and 'user' types should have venue assignments
        const venueAssignments = (userType === 'admin' || userType === 'user') ? (updatedRow[10] || 'ALL') : '';
        
        return {
          success: true,
          message: 'Profile updated successfully',
          user: {
            username: updatedRow[0] || '',
            fullName: updatedRow[2] || '',
            department: updatedRow[3] || '',
            email: updatedRow[4] || '',
            phone: (updatedRow[5] && updatedRow[5].toString() !== '#ERROR!') ? updatedRow[5].toString() : '',
            userType: userType,
            status: updatedRow[7] || 'Active',
            profilePicture: updatedRow[8] || '',
            venueAssignments: venueAssignments
          }
        };
      }
    }
    
    return { success: false, error: 'User not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Send message
function sendMessage(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(MESSAGES_SHEET_NAME);
    
    if (!sheet) {
      // Create Messages sheet with headers if it doesn't exist
      sheet = ss.insertSheet(MESSAGES_SHEET_NAME);
      sheet.appendRow(['Message ID', 'Sender Username', 'Receiver Username', 'Message', 'Timestamp', 'Read Status']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 6);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 150); // Message ID
      sheet.setColumnWidth(2, 150); // Sender Username
      sheet.setColumnWidth(3, 150); // Receiver Username
      sheet.setColumnWidth(4, 300); // Message
      sheet.setColumnWidth(5, 150); // Timestamp
      sheet.setColumnWidth(6, 100); // Read Status
    }
    
    const messageId = 'MSG-' + Date.now();
    const timestamp = new Date();
    
    sheet.appendRow([
      messageId,
      data.senderUsername,
      data.receiverUsername,
      data.message,
      timestamp,
      'unread'
    ]);
    
    return {
      success: true,
      message: 'Message sent successfully',
      messageId: messageId,
      timestamp: timestamp.toISOString()
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Get messages between two users
function getMessages(user1, user2) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(MESSAGES_SHEET_NAME);
    
    if (!sheet) {
      return { success: true, messages: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const messages = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Message ID exists
        const senderUsername = row[1];
        const receiverUsername = row[2];
        
        // Check if message is between the two users
        if ((senderUsername === user1 && receiverUsername === user2) ||
            (senderUsername === user2 && receiverUsername === user1)) {
          messages.push({
            messageId: row[0],
            senderUsername: senderUsername,
            receiverUsername: receiverUsername,
            message: row[3],
            timestamp: row[4],
            readStatus: row[5] || 'unread'
          });
        }
      }
    }
    
    // Sort messages by timestamp (oldest first)
    messages.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
    
    return {
      success: true,
      messages: messages
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Get all users (for messaging interface)
function getAllUsers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: true, users: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const users = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[7] === 'Active') { // Username exists and user is active (Status is now column 7)
        const userType = row[6]; // Get actual user type from sheet without default fallback
        // Only 'admin' and 'user' types should have venue assignments
        const venueAssignments = (userType === 'admin' || userType === 'user') ? (row[10] || 'ALL') : '';
        
        users.push({
          username: row[0] || '',
          fullName: row[2] || '',
          department: row[3] || '',
          email: row[4] || '',
          phone: (row[5] && row[5].toString() !== '#ERROR!') ? row[5].toString() : '',
          userType: userType,
          status: row[7] || 'Active',
          profilePicture: row[8] || '',
          venueAssignments: venueAssignments
        });
      }
    }
    
    return {
      success: true,
      users: users
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Mark messages as read
function markMessagesAsRead(currentUser, otherUser) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(MESSAGES_SHEET_NAME);
    
    if (!sheet) {
      return { success: true, message: 'No messages to mark as read' };
    }
    
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Message ID exists
        const senderUsername = row[1];
        const receiverUsername = row[2];
        const readStatus = row[5];
        
        // Mark messages as read if they were sent TO the current user FROM the other user
        // and they are currently unread
        if (senderUsername === otherUser && 
            receiverUsername === currentUser && 
            readStatus === 'unread') {
          sheet.getRange(i + 1, 6).setValue('read'); // Column 6 is Read Status
          updatedCount++;
        }
      }
    }
    
    return {
      success: true,
      message: `Marked ${updatedCount} messages as read`,
      updatedCount: updatedCount
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Get unread message counts for each user
function getUnreadMessageCounts(currentUser) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(MESSAGES_SHEET_NAME);
    
    if (!sheet) {
      return { success: true, unreadCounts: {} };
    }
    
    const data = sheet.getDataRange().getValues();
    const unreadCounts = {};
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Message ID exists
        const senderUsername = row[1];
        const receiverUsername = row[2];
        const readStatus = row[5];
        
        // Count unread messages sent TO the current user
        if (receiverUsername === currentUser && readStatus === 'unread') {
          if (!unreadCounts[senderUsername]) {
            unreadCounts[senderUsername] = 0;
          }
          unreadCounts[senderUsername]++;
        }
      }
    }
    
    return {
      success: true,
      unreadCounts: unreadCounts
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Add usage log
function addUsageLog(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!sheet) {
      // Create Logs sheet with headers if it doesn't exist
      sheet = ss.insertSheet(LOGS_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Log ID', 'Venue', 'Activity Date', 'Day', 'Department', 'Start Time', 'End Time', 'Date Submitted', 'Purpose', 'Facilities']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 11);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths and formats
      sheet.setColumnWidth(1, 150); // Timestamp
      sheet.setColumnWidth(2, 120); // Log ID
      sheet.setColumnWidth(3, 150); // Venue
      sheet.setColumnWidth(4, 120); // Activity Date
      sheet.setColumnWidth(5, 100); // Day
      sheet.setColumnWidth(6, 150); // Department
      sheet.setColumnWidth(7, 120); // Start Time
      sheet.setColumnWidth(8, 120); // End Time
      sheet.setColumnWidth(9, 150); // Date Submitted
      sheet.setColumnWidth(10, 200); // Purpose
      sheet.setColumnWidth(11, 200); // Facilities
      
      // Format time columns as text to prevent date conversion
      const startTimeRange = sheet.getRange('G:G'); // Column G (Start Time)
      const endTimeRange = sheet.getRange('H:H');   // Column H (End Time)
      startTimeRange.setNumberFormat('@'); // @ means text format
      endTimeRange.setNumberFormat('@');   // @ means text format
    } else {
      // If sheet exists, ensure time columns are formatted as text
      const startTimeRange = sheet.getRange('G:G'); // Column G (Start Time)
      const endTimeRange = sheet.getRange('H:H');   // Column H (End Time)
      startTimeRange.setNumberFormat('@'); // @ means text format
      endTimeRange.setNumberFormat('@');   // @ means text format
    }
    
    const logId = 'LOG-' + Date.now();
    const timestamp = new Date();
    
    // Parse and format the activity date properly with extra validation
    let activityDate = data.activityDate;
    
    // Log what we received for debugging
    console.log('Received activityDate:', activityDate, 'Type:', typeof activityDate);
    
    if (activityDate) {
      // If it's a date string in YYYY-MM-DD format, keep it as string to avoid timezone issues
      if (typeof activityDate === 'string' && activityDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
        // Keep as string - don't convert to Date object to avoid timezone problems
        console.log('Keeping date as string to avoid timezone issues:', activityDate);
        // Optionally validate that it's a reasonable date
        const [year, month, day] = activityDate.split('-');
        const yearNum = parseInt(year);
        const monthNum = parseInt(month);
        const dayNum = parseInt(day);
        
        if (yearNum < 1900 || yearNum > 2100 || monthNum < 1 || monthNum > 12 || dayNum < 1 || dayNum > 31) {
          console.log('Invalid date detected:', activityDate, 'Setting to null');
          activityDate = null;
        }
      } else if (typeof activityDate === 'string') {
        // Handle other date formats by converting to Date and then back to YYYY-MM-DD string
        const parsedDate = new Date(activityDate);
        if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() >= 1900 && parsedDate.getFullYear() <= 2100) {
          // Convert back to YYYY-MM-DD string to maintain consistency
          const year = parsedDate.getFullYear();
          const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
          const day = String(parsedDate.getDate()).padStart(2, '0');
          activityDate = `${year}-${month}-${day}`;
          console.log('Converted other format to YYYY-MM-DD string:', activityDate);
        } else {
          console.log('Could not parse date string:', activityDate, 'Setting to null');
          activityDate = null;
        }
      } else if (activityDate instanceof Date) {
        // If it's already a Date object, convert to YYYY-MM-DD string
        if (!isNaN(activityDate.getTime()) && activityDate.getFullYear() >= 1900 && activityDate.getFullYear() <= 2100) {
          const year = activityDate.getFullYear();
          const month = String(activityDate.getMonth() + 1).padStart(2, '0');
          const day = String(activityDate.getDate()).padStart(2, '0');
          activityDate = `${year}-${month}-${day}`;
          console.log('Converted Date object to YYYY-MM-DD string:', activityDate);
        } else {
          console.log('Invalid Date object:', activityDate, 'Setting to null');
          activityDate = null;
        }
      }
    }
    
    // Parse dateSubmitted if provided, otherwise use current timestamp
    let dateSubmitted = data.dateSubmitted ? new Date(data.dateSubmitted) : timestamp;
    if (isNaN(dateSubmitted.getTime())) {
      dateSubmitted = timestamp;
    }
    
    // Handle start and end time properly - keep as strings to avoid Google Sheets date conversion
    let startTime = data.startTime || '';
    let endTime = data.endTime || '';
    
    // Log what we received for debugging
    console.log('Received startTime:', startTime, 'Type:', typeof startTime);
    console.log('Received endTime:', endTime, 'Type:', typeof endTime);
    
    // Clean and validate time formats
    if (startTime) {
      // If it's already a proper time string (HH:MM format), convert to 12-hour format
      if (typeof startTime === 'string' && startTime.match(/^\d{1,2}:\d{2}$/)) {
        const [hours24Str, minutes] = startTime.split(':');
        const hours24 = parseInt(hours24Str);
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
        startTime = `${hours12}:${minutes} ${ampm}`;
      } else if (startTime instanceof Date || typeof startTime === 'object') {
        // If it's a Date object (shouldn't happen, but just in case)
        const hours24 = startTime.getHours();
        const minutes = startTime.getMinutes().toString().padStart(2, '0');
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
        startTime = `${hours12}:${minutes} ${ampm}`;
      } else if (typeof startTime === 'string' && (startTime.includes('AM') || startTime.includes('PM'))) {
        // Already in 12-hour format, keep as is
        startTime = startTime;
      } else {
        // Convert to string and try to extract time
        startTime = startTime.toString();
      }
    }
    
    if (endTime) {
      // If it's already a proper time string (HH:MM format), convert to 12-hour format
      if (typeof endTime === 'string' && endTime.match(/^\d{1,2}:\d{2}$/)) {
        const [hours24Str, minutes] = endTime.split(':');
        const hours24 = parseInt(hours24Str);
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
        endTime = `${hours12}:${minutes} ${ampm}`;
      } else if (endTime instanceof Date || typeof endTime === 'object') {
        // If it's a Date object (shouldn't happen, but just in case)
        const hours24 = endTime.getHours();
        const minutes = endTime.getMinutes().toString().padStart(2, '0');
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
        endTime = `${hours12}:${minutes} ${ampm}`;
      } else if (typeof endTime === 'string' && (endTime.includes('AM') || endTime.includes('PM'))) {
        // Already in 12-hour format, keep as is
        endTime = endTime;
      } else {
        // Convert to string and try to extract time
        endTime = endTime.toString();
      }
    }
    
    console.log('Final activityDate before saving:', activityDate);
    console.log('Final startTime before saving:', startTime);
    console.log('Final endTime before saving:', endTime);
    
    // Get the row number before inserting
    const nextRow = sheet.getLastRow() + 1;
    
    // First, set the time columns to text format for the row we're about to insert
    sheet.getRange(nextRow, 7).setNumberFormat('@'); // Start Time column
    sheet.getRange(nextRow, 8).setNumberFormat('@'); // End Time column
    
    // Now append the row data
    sheet.appendRow([
      timestamp,
      logId,
      data.venue,
      activityDate,
      data.day || '',
      data.department,
      startTime,
      endTime,
      dateSubmitted,
      data.purpose || '',
      data.facilities || ''
    ]);
    
    // Double-check: explicitly set the time values as text using setValues
    const lastRow = sheet.getLastRow();
    if (startTime) {
      sheet.getRange(lastRow, 7).setValue(startTime.toString()); // Force as string
    }
    if (endTime) {
      sheet.getRange(lastRow, 8).setValue(endTime.toString()); // Force as string
    }
    
    return {
      success: true,
      message: 'Usage log added successfully',
      logId: logId,
      timestamp: timestamp.toISOString()
    };
  } catch (error) {
    console.log('Error in addUsageLog:', error);
    return { success: false, error: error.toString() };
  }
}

// Get usage logs
function getUsageLogs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!sheet) {
      return { success: true, usageLogs: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const usageLogs = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Timestamp exists
        // Handle date values properly
        let timestamp = row[0];
        let activityDate = row[3];
        let dateSubmitted = row[8];
        
        // Convert timestamps to proper Date objects if they're not already
        if (timestamp && !(timestamp instanceof Date)) {
          timestamp = new Date(timestamp);
        }
        
        // Handle activity date - keep as string if it's in YYYY-MM-DD format
        if (activityDate) {
          if (typeof activityDate === 'string' && activityDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
            // Already in correct YYYY-MM-DD format, keep as string
            console.log('Activity date is already in YYYY-MM-DD format:', activityDate);
          } else if (activityDate instanceof Date) {
            // Convert Date object to YYYY-MM-DD string
            if (activityDate.getFullYear() >= 1900 && activityDate.getFullYear() <= 2100) {
              const year = activityDate.getFullYear();
              const month = String(activityDate.getMonth() + 1).padStart(2, '0');
              const day = String(activityDate.getDate()).padStart(2, '0');
              activityDate = `${year}-${month}-${day}`;
              console.log('Converted activity date from Date to YYYY-MM-DD:', activityDate);
            } else {
              console.log('Invalid activity date year:', activityDate.getFullYear(), 'Setting to null');
              activityDate = null;
            }
          } else if (typeof activityDate === 'number') {
            // Handle Google Sheets serial number (if any)
            if (activityDate > 25569) { // Roughly 1970-01-01 in Excel serial date
              const dateObj = new Date((activityDate - 25569) * 86400 * 1000);
              if (dateObj.getFullYear() >= 1900 && dateObj.getFullYear() <= 2100) {
                const year = dateObj.getFullYear();
                const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                const day = String(dateObj.getDate()).padStart(2, '0');
                activityDate = `${year}-${month}-${day}`;
                console.log('Converted serial number to YYYY-MM-DD:', activityDate);
              } else {
                activityDate = null;
              }
            } else {
              activityDate = null;
            }
          } else {
            // Try to parse as string and convert to YYYY-MM-DD
            const dateObj = new Date(activityDate);
            if (!isNaN(dateObj.getTime()) && dateObj.getFullYear() >= 1900 && dateObj.getFullYear() <= 2100) {
              const year = dateObj.getFullYear();
              const month = String(dateObj.getMonth() + 1).padStart(2, '0');
              const day = String(dateObj.getDate()).padStart(2, '0');
              activityDate = `${year}-${month}-${day}`;
              console.log('Parsed and converted to YYYY-MM-DD:', activityDate);
            } else {
              console.log('Could not parse activity date:', activityDate, 'Setting to null');
              activityDate = null;
            }
          }
        }
        
        if (dateSubmitted && !(dateSubmitted instanceof Date)) {
          dateSubmitted = new Date(dateSubmitted);
          // Check for problematic dates
          if (isNaN(dateSubmitted.getTime()) || dateSubmitted.getFullYear() < 1900) {
            dateSubmitted = null;
          }
        }
        
        // Handle time values properly - they should be strings, not dates
        let startTime = row[6];
        let endTime = row[7];
        
        console.log('Retrieved startTime from sheet:', startTime, 'Type:', typeof startTime, 'Is Date:', startTime instanceof Date);
        console.log('Retrieved endTime from sheet:', endTime, 'Type:', typeof endTime, 'Is Date:', endTime instanceof Date);
        
        // Process startTime
        if (startTime instanceof Date) {
          console.log('StartTime is Date, year:', startTime.getFullYear());
          // Check if it's the problematic 1899 date (indicates time was converted to date)
          if (startTime.getFullYear() === 1899) {
            // Extract time and convert to 12-hour format
            const hours24 = startTime.getHours();
            const minutes = startTime.getMinutes().toString().padStart(2, '0');
            const ampm = hours24 >= 12 ? 'PM' : 'AM';
            const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
            startTime = `${hours12}:${minutes} ${ampm}`;
            console.log('Converted 1899 date to 12-hour time:', startTime);
          } else {
            // If it's a real date, extract time part and convert to 12-hour format
            const hours24 = startTime.getHours();
            const minutes = startTime.getMinutes().toString().padStart(2, '0');
            const ampm = hours24 >= 12 ? 'PM' : 'AM';
            const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
            startTime = `${hours12}:${minutes} ${ampm}`;
            console.log('Extracted time from real date to 12-hour format:', startTime);
          }
        } else if (startTime && typeof startTime === 'string') {
          // If it's already a string, validate and clean it
          if (startTime.includes('1899-12-30T')) {
            // Extract time from ISO string and convert to 12-hour format
            const timeMatch = startTime.match(/T(\d{2}):(\d{2}):/);
            if (timeMatch) {
              const hours24 = parseInt(timeMatch[1]);
              const minutes = timeMatch[2];
              const ampm = hours24 >= 12 ? 'PM' : 'AM';
              const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
              startTime = `${hours12}:${minutes} ${ampm}`;
              console.log('Extracted time from ISO string to 12-hour format:', startTime);
            }
          } else if (startTime.match(/^\d{1,2}:\d{2}$/)) {
            // Already in HH:MM format, convert to 12-hour format 
            const [hours24Str, minutes] = startTime.split(':');
            const hours24 = parseInt(hours24Str);
            const ampm = hours24 >= 12 ? 'PM' : 'AM';
            const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
            startTime = `${hours12}:${minutes} ${ampm}`;
          } else if (startTime.includes('AM') || startTime.includes('PM')) {
            // Already in 12-hour format, keep as is
            startTime = startTime;
          }
        } else if (startTime) {
          startTime = startTime.toString();
        } else {
          startTime = '';
        }
        
        // Process endTime
        if (endTime instanceof Date) {
          console.log('EndTime is Date, year:', endTime.getFullYear());
          // Check if it's the problematic 1899 date (indicates time was converted to date)
          if (endTime.getFullYear() === 1899) {
            // Extract time and convert to 12-hour format
            const hours24 = endTime.getHours();
            const minutes = endTime.getMinutes().toString().padStart(2, '0');
            const ampm = hours24 >= 12 ? 'PM' : 'AM';
            const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
            endTime = `${hours12}:${minutes} ${ampm}`;
            console.log('Converted 1899 date to 12-hour time:', endTime);
          } else {
            // If it's a real date, extract time part and convert to 12-hour format
            const hours24 = endTime.getHours();
            const minutes = endTime.getMinutes().toString().padStart(2, '0');
            const ampm = hours24 >= 12 ? 'PM' : 'AM';
            const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
            endTime = `${hours12}:${minutes} ${ampm}`;
            console.log('Extracted time from real date to 12-hour format:', endTime);
          }
        } else if (endTime && typeof endTime === 'string') {
          // If it's already a string, validate and clean it
          if (endTime.includes('1899-12-30T')) {
            // Extract time from ISO string and convert to 12-hour format
            const timeMatch = endTime.match(/T(\d{2}):(\d{2}):/);
            if (timeMatch) {
              const hours24 = parseInt(timeMatch[1]);
              const minutes = timeMatch[2];
              const ampm = hours24 >= 12 ? 'PM' : 'AM';
              const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
              endTime = `${hours12}:${minutes} ${ampm}`;
              console.log('Extracted time from ISO string to 12-hour format:', endTime);
            }
          } else if (endTime.match(/^\d{1,2}:\d{2}$/)) {
            // Already in HH:MM format, convert to 12-hour format
            const [hours24Str, minutes] = endTime.split(':');
            const hours24 = parseInt(hours24Str);
            const ampm = hours24 >= 12 ? 'PM' : 'AM';
            const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
            endTime = `${hours12}:${minutes} ${ampm}`;
          } else if (endTime.includes('AM') || endTime.includes('PM')) {
            // Already in 12-hour format, keep as is
            endTime = endTime;
          }
        } else if (endTime) {
          endTime = endTime.toString();
        } else {
          endTime = '';
        }
        
        console.log('Final startTime:', startTime, 'Final endTime:', endTime);
        
        usageLogs.push({
          timestamp: timestamp ? timestamp.toISOString() : null,
          logId: row[1],
          venue: row[2],
          activityDate: activityDate, // Already in YYYY-MM-DD format as string
          day: row[4],
          department: row[5],
          startTime: startTime,
          endTime: endTime,
          dateSubmitted: dateSubmitted ? dateSubmitted.toISOString() : null,
          purpose: row[9],
          facilities: row[10]
        });
      }
    }
    
    // Sort by timestamp (most recent first)
    usageLogs.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return {
      success: true,
      usageLogs: usageLogs
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Update usage log
function updateUsageLog(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Usage logs sheet not found' };
    }
    
    const sheetData = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Find the row with the matching log ID
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][1] && sheetData[i][1].toString() === data.logId.toString()) {
        rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Usage log not found' };
    }
    
    // Parse and format the activity date properly
    let activityDate = data.activityDate;
    if (activityDate) {
      // If it's a date string in YYYY-MM-DD format, keep it as string to avoid timezone issues
      if (typeof activityDate === 'string' && activityDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
        // Keep as string - don't convert to Date object to avoid timezone problems
        console.log('Keeping date as string to avoid timezone issues:', activityDate);
      } else if (typeof activityDate === 'string') {
        // Handle other date formats by converting to Date and then back to YYYY-MM-DD string
        const parsedDate = new Date(activityDate);
        if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() >= 1900 && parsedDate.getFullYear() <= 2100) {
          // Convert back to YYYY-MM-DD string to maintain consistency
          const year = parsedDate.getFullYear();
          const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
          const day = String(parsedDate.getDate()).padStart(2, '0');
          activityDate = `${year}-${month}-${day}`;
          console.log('Converted other format to YYYY-MM-DD string:', activityDate);
        } else {
          console.log('Could not parse date string:', activityDate, 'Setting to null');
          activityDate = null;
        }
      } else if (activityDate instanceof Date) {
        // If it's already a Date object, convert to YYYY-MM-DD string
        if (!isNaN(activityDate.getTime()) && activityDate.getFullYear() >= 1900 && activityDate.getFullYear() <= 2100) {
          const year = activityDate.getFullYear();
          const month = String(activityDate.getMonth() + 1).padStart(2, '0');
          const day = String(activityDate.getDate()).padStart(2, '0');
          activityDate = `${year}-${month}-${day}`;
          console.log('Converted Date object to YYYY-MM-DD string:', activityDate);
        } else {
          console.log('Invalid Date object:', activityDate, 'Setting to null');
          activityDate = null;
        }
      }
    }
    
    // Handle start and end time properly - convert from 24-hour to 12-hour format
    let startTime = data.startTime || '';
    let endTime = data.endTime || '';
    
    if (startTime) {
      // If it's already a proper time string (HH:MM format), convert to 12-hour format
      if (typeof startTime === 'string' && startTime.match(/^\d{1,2}:\d{2}$/)) {
        const [hours24Str, minutes] = startTime.split(':');
        const hours24 = parseInt(hours24Str);
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
        startTime = `${hours12}:${minutes} ${ampm}`;
      }
    }
    
    if (endTime) {
      // If it's already a proper time string (HH:MM format), convert to 12-hour format
      if (typeof endTime === 'string' && endTime.match(/^\d{1,2}:\d{2}$/)) {
        const [hours24Str, minutes] = endTime.split(':');
        const hours24 = parseInt(hours24Str);
        const ampm = hours24 >= 12 ? 'PM' : 'AM';
        const hours12 = hours24 % 12 || 12; // Convert to 12-hour format, 0 becomes 12
        endTime = `${hours12}:${minutes} ${ampm}`;
      }
    }
    
    // Update the row data (keeping the original timestamp and log ID)
    const originalTimestamp = sheetData[rowIndex - 1][0]; // Keep original timestamp
    const originalDateSubmitted = sheetData[rowIndex - 1][8]; // Keep original date submitted
    
    // Set the time columns to text format before updating
    sheet.getRange(rowIndex, 7).setNumberFormat('@'); // Start Time column
    sheet.getRange(rowIndex, 8).setNumberFormat('@'); // End Time column
    
    // Update the row
    sheet.getRange(rowIndex, 1, 1, 11).setValues([[
      originalTimestamp,
      data.logId,
      data.venue,
      activityDate,
      data.day || '',
      data.department,
      startTime,
      endTime,
      originalDateSubmitted,
      data.purpose,
      data.facilities || ''
    ]]);
    
    // Double-check: explicitly set the time values as text using setValues
    if (startTime) {
      sheet.getRange(rowIndex, 7).setValue(startTime.toString()); // Force as string
    }
    if (endTime) {
      sheet.getRange(rowIndex, 8).setValue(endTime.toString()); // Force as string
    }
    
    return {
      success: true,
      message: 'Usage log updated successfully',
      logId: data.logId
    };
  } catch (error) {
    console.log('Error in updateUsageLog:', error);
    return { success: false, error: error.toString() };
  }
}

// Get all users for management (admin only)
function getUsers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: true, users: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const users = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Username exists
        const userType = row[6]; // Get actual user type from sheet without default fallback
        // Only 'admin' and 'user' types should have venue assignments
        const venueAssignments = (userType === 'admin' || userType === 'user') ? (row[10] || 'ALL') : '';
        
        users.push({
          username: row[0] || '',
          fullName: row[2] || '',
          department: row[3] || '',
          email: row[4] || '',
          phone: (row[5] && row[5].toString() !== '#ERROR!') ? row[5].toString() : '',
          userType: userType,
          status: row[7] || 'Active',
          profilePicture: row[8] || '',
          createdDate: row[9] || '',
          venueAssignments: venueAssignments
        });
      }
    }
    
    return {
      success: true,
      users: users
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Update user status (admin only)
function updateUserStatus(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    const sheetData = sheet.getDataRange().getValues();
    
    // Find the user row
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      if (row[0] && row[0].toString().toLowerCase() === data.username.toLowerCase()) {
        // Update user status
        sheet.getRange(i + 1, 8).setValue(data.status); // Column 8 is Status (was 7, now 8 due to department)
        
        return {
          success: true,
          message: `User ${data.username} status updated to ${data.status}`
        };
      }
    }
    
    return { success: false, error: 'User not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Delete user (admin only)
function deleteUser(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    const sheetData = sheet.getDataRange().getValues();
    
    // Find the user row
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      if (row[0] && row[0].toString().toLowerCase() === username.toLowerCase()) {
        // Delete the user row
        sheet.deleteRow(i + 1);
        
        return {
          success: true,
          message: `User ${username} deleted successfully`
        };
      }
    }
    
    return { success: false, error: 'User not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Update user venue assignments (admin only)
function updateUserVenueAssignments(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }
    
    const sheetData = sheet.getDataRange().getValues();
    
    // Find the user row
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      if (row[0] && row[0].toString().toLowerCase() === data.username.toLowerCase()) {
        // Check user type first - only 'admin' and 'user' types can have venue assignments
        const userType = row[6]; // Get actual user type from sheet without default fallback
        
        if (userType === 'admin' || userType === 'user') {
          // Update venue assignments - Column 11 (index 10)
          const venueAssignments = data.venueAssignments || 'ALL';
          sheet.getRange(i + 1, 11).setValue(venueAssignments);
          
          return {
            success: true,
            message: `Venue assignments updated for user ${data.username}`
          };
        } else {
          // For non-admin, non-user types, clear venue assignments
          sheet.getRange(i + 1, 11).setValue('');
          
          return {
            success: false,
            error: `User type '${userType}' cannot have venue assignments. Venue assignments cleared.`
          };
        }
      }
    }
    
    return { success: false, error: 'User not found' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ====== JOB ORDER FUNCTIONALITY ======

// Job Order Sheet Configuration
const JOB_ORDERS_SHEET_NAME = 'JobOrders';

// Submit job order - generates form number and saves to sheet
function submitJobOrder(data) {
  try {
    // Log the formData to ensure it's structured correctly
    Logger.log("Received job order data: " + JSON.stringify(data));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(JOB_ORDERS_SHEET_NAME);
    
    if (!sheet) {
      // Create the JobOrders sheet if it doesn't exist
      sheet = ss.insertSheet(JOB_ORDERS_SHEET_NAME);
      sheet.appendRow(['Form Number', 'Date Filed', 'Department', 'Location', 'Nature of Work', 'Requested By', 'Status', 'Date Submitted']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 8);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 120); // Form Number
      sheet.setColumnWidth(2, 120); // Date Filed
      sheet.setColumnWidth(3, 200); // Department
      sheet.setColumnWidth(4, 200); // Location
      sheet.setColumnWidth(5, 250); // Nature of Work
      sheet.setColumnWidth(6, 150); // Requested By
      sheet.setColumnWidth(7, 100); // Status
      sheet.setColumnWidth(8, 150); // Date Submitted
    }

    // Generate the next form number
    const nextFormNumber = generateNextFormNumber(sheet);
    
    // Prepare the data
    const timestamp = new Date();
    const dateFiled = data.dateFiled ? new Date(data.dateFiled) : timestamp;
    
    // Format date to avoid timezone issues
    const formattedDateFiled = Utilities.formatDate(dateFiled, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Add the new job order
    sheet.appendRow([
      nextFormNumber,
      formattedDateFiled,
      data.department,
      data.location,
      data.natureOfWork,
      data.requestedBy,
      'Pending',
      timestamp
    ]);
    
    // Generate print content
    const printContent = generateJobOrderPrintContent({
      formNumber: nextFormNumber,
      dateFiled: formattedDateFiled,
      department: data.department,
      location: data.location,
      natureOfWork: data.natureOfWork,
      requestedBy: data.requestedBy
    });
    
    return {
      success: true,
      message: 'Job order submitted successfully',
      formNumber: nextFormNumber,
      printContent: printContent
    };

  } catch (error) {
    Logger.log(`Error in submitJobOrder: ${error.message}`);
    return { success: false, error: error.toString() };
  }
}

// ====== AIRCON CONCERNS FUNCTIONALITY ======

// Submit aircon concern - saves to AC-Concerns sheet
function submitAirconConcern(data) {
  try {
    Logger.log("Received aircon concern data: " + JSON.stringify(data));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(AC_CONCERNS_SHEET_NAME);
    
    if (!sheet) {
      // Create the AC-Concerns sheet if it doesn't exist
      sheet = ss.insertSheet(AC_CONCERNS_SHEET_NAME);
      sheet.appendRow(['DATE', 'Department', 'Aircon Type', 'Concern/s', 'Reported by', 'Status', 'Date Accomplished', 'Date Submitted']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 8);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 120); // DATE
      sheet.setColumnWidth(2, 150); // Department
      sheet.setColumnWidth(3, 120); // Aircon Type
      sheet.setColumnWidth(4, 300); // Concern/s
      sheet.setColumnWidth(5, 150); // Reported by
      sheet.setColumnWidth(6, 100); // Status
      sheet.setColumnWidth(7, 150); // Date Accomplished
      sheet.setColumnWidth(8, 150); // Date Submitted
    }

    // Prepare the data
    const timestamp = new Date();
    const concernDate = data.date ? new Date(data.date) : timestamp;
    
    // Format date to avoid timezone issues
    const formattedConcernDate = Utilities.formatDate(concernDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Add the new aircon concern
    sheet.appendRow([
      formattedConcernDate,
      data.department,
      data.airconType,
      data.concerns,
      data.reportedBy,
      data.status || 'Pending',
      data.dateAccomplished || '', // Initially empty
      timestamp
    ]);
    
    return {
      success: true,
      message: 'Aircon concern submitted successfully'
    };

  } catch (error) {
    Logger.log(`Error in submitAirconConcern: ${error.message}`);
    return { success: false, error: error.toString() };
  }
}

// Get AC concerns for a specific department or all concerns (for admin)
function getACConcerns(department) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(AC_CONCERNS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'AC Concerns sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const concerns = [];
    
    // Skip header row (index 0) and process data
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const concern = {
        id: `${row[0]}-${row[4]}`, // Generate ID from date and reporter
        date: row[0],
        department: row[1],
        airconType: row[2],
        concerns: row[3],
        reportedBy: row[4],
        status: row[5],
        dateAccomplished: row[6],
        dateSubmitted: row[7]
      };
      
      // Filter by department if specified (for non-admin users)
      if (!department || concern.department === department) {
        concerns.push(concern);
      }
    }
    
    // Sort by date submitted (newest first)
    concerns.sort((a, b) => new Date(b.dateSubmitted) - new Date(a.dateSubmitted));
    
    return {
      success: true,
      concerns: concerns
    };
  } catch (error) {
    Logger.log(`Error in getACConcerns: ${error.message}`);
    return { success: false, error: error.toString() };
  }
}

// Update AC concern details
function updateACConcern(concernId, concerns, status, dateAccomplished) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(AC_CONCERNS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'AC Concerns sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    
    // Find the concern to update
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const currentId = `${row[0]}-${row[4]}`; // Generate ID from date and reporter
      
      if (currentId === concernId) {
        // Update the concern details
        if (concerns) {
          sheet.getRange(i + 1, 4).setValue(concerns); // Column D: Concern/s
        }
        if (status) {
          sheet.getRange(i + 1, 6).setValue(status); // Column F: Status
          
          // Automatically set date accomplished when status is set to "Completed"
          if (status === 'Completed' && !dateAccomplished) {
            const today = new Date();
            const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            dateAccomplished = formattedDate;
          }
        }
        if (dateAccomplished) {
          sheet.getRange(i + 1, 7).setValue(dateAccomplished); // Column G: Date Accomplished
        }
        
        return {
          success: true,
          message: 'AC concern updated successfully'
        };
      }
    }
    
    return { success: false, error: 'AC concern not found' };
  } catch (error) {
    Logger.log(`Error in updateACConcern: ${error.message}`);
    return { success: false, error: error.toString() };
  }
}

// Get job orders for a specific user or all job orders (for admin)
function getJobOrders(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(JOB_ORDERS_SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Job Orders sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const jobOrders = [];
    
    if (data.length <= 1) {
      // Only headers or empty sheet
      return {
        success: true,
        jobOrders: []
      };
    }
    
    // Start from row 2 to skip headers
    // Headers: [Form Number, Date Filed, Department, Location, Nature of Work, Requested By, Status, Date Submitted]
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log(`DEBUG: Processing row ${i}: ${JSON.stringify(row)}`);
      
      // If username is provided and user is not admin, filter by requested by
      // For now, we'll show all orders to the user (you can modify this logic)
      const jobOrder = {
        formNumber: row[0],
        dateFiled: row[1],
        department: row[2], 
        location: row[3],
        natureOfWork: row[4],
        requestedBy: row[5],
        status: row[6] || 'Pending',
        dateSubmitted: row[7],
        printContent: generateJobOrderPrintContent({
          formNumber: row[0],
          department: row[2],
          dateFiled: row[1],
          location: row[3],
          natureOfWork: row[4],
          requestedBy: row[5]
        })
      };
      
      Logger.log(`DEBUG: Created jobOrder object: ${JSON.stringify(jobOrder)}`);
      Logger.log(`DEBUG: Department value in jobOrder: "${jobOrder.department}"`);
      
      // If username is provided, filter by requested by (unless admin)
      if (username && !isAdmin(username)) {
        // Get user info to check if they requested this job order
        const userInfo = getUserInfo(username);
        if (userInfo && userInfo.fullName === row[5]) {
          jobOrders.push(jobOrder);
        }
      } else {
        // Show all job orders (for admin or when no username filter)
        jobOrders.push(jobOrder);
      }
    }
    
    // Sort by date submitted (newest first)
    jobOrders.sort((a, b) => new Date(b.dateSubmitted) - new Date(a.dateSubmitted));
    
    return {
      success: true,
      jobOrders: jobOrders
    };
    
  } catch (error) {
    Logger.log(`Error in getJobOrders: ${error.message}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Update job order status (admin only)
 */
function updateJobOrder(formNumber, status) {
  try {
    Logger.log(`updateJobOrder called with: formNumber=${formNumber}, status=${status}`);
    Logger.log(`DEBUG: Using sheet name: ${JOB_ORDERS_SHEET_NAME}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEBUG: List all available sheets
    const allSheets = ss.getSheets();
    const sheetNames = allSheets.map(sheet => sheet.getName());
    Logger.log(`DEBUG: Available sheets: ${JSON.stringify(sheetNames)}`);
    
    let sheet = ss.getSheetByName(JOB_ORDERS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('DEBUG: Job orders sheet not found - this should be the error message');
      return { success: false, error: 'Job orders sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log(`Sheet has ${data.length} rows`);
    
    // Find the job order by form number (in column A, index 0)
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      Logger.log(`Row ${i}: formNumber=${data[i][0]}, comparing with ${formNumber}`);
      if (data[i][0] === formNumber) { // Form number is in column A (index 0)
        rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        Logger.log(`Found job order at row ${rowIndex}`);
        break;
      }
    }
    
    if (rowIndex === -1) {
      Logger.log(`Job order with form number ${formNumber} not found`);
      return { success: false, error: 'Job order not found' };
    }
    
    // Update status (column G, index 6)
    Logger.log(`Updating row ${rowIndex}: status in column G`);
    sheet.getRange(rowIndex, 7).setValue(status); // Column G - Status (1-indexed, so column 7)
    
    // Log the update
    Logger.log(`Successfully updated job order ${formNumber}`);
    Logger.log(`UPDATE: Job order ${formNumber} status updated to: ${status}`);
    
    return {
      success: true,
      message: 'Job order status updated successfully'
    };
    
  } catch (error) {
    Logger.log(`Error in updateJobOrder: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    return { success: false, error: error.toString() };
  }
}

// Helper function to check if user is admin
function isAdmin(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return false;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === username) { // Username is in column A
        return row[6] === 'admin'; // User Type is in column G (index 6)
      }
    }
    
    return false;
  } catch (error) {
    return false;
  }
}

// Helper function to get user info
function getUserInfo(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === username) { // Username is in column A
        return {
          username: row[0],
          fullName: row[2],
          department: row[3],
          email: row[4],
          userType: row[6]
        };
      }
    }
    
    return null;
  } catch (error) {
    return null;
  }
}

// Generate the next form number in sequence
function generateNextFormNumber(sheet) {
  const data = sheet.getDataRange().getValues();
  let maxNumber = 0;

  // Start from row 2 to skip headers
  for (let i = 1; i < data.length; i++) {
    const formNumber = data[i][0];
    if (formNumber && typeof formNumber === 'string') {
      const match = formNumber.match(/^B-(\d+)$/);
      if (match) {
        maxNumber = Math.max(maxNumber, parseInt(match[1], 10));
      }
    }
  }

  return `B-${maxNumber + 1}`;
}

// Generate print content for job order
function generateJobOrderPrintContent(formData) {
  // Add logging to debug
  Logger.log("Generating print content with data: " + JSON.stringify(formData));
  
  return `
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>pao_basak2025</title>
    <style>
      @page {
        size: letter; /* Set the page size to letter */
        margin: 0.25in; /* Set the margins to minimum */
      }

      body {
        font-family: 'Arial', sans-serif;
        margin: 0;
        padding: 0;
        display: flex;
        justify-content: center;
        align-items: flex-start; /* Align items to the top */
        min-height: 100vh;
        background: linear-gradient(135deg, #f5f7fa, #c3cfe2);
      }

      .wrapper {
        width: 100%;
        max-width: 1200px;
        padding: 0; /* Remove padding */
        box-sizing: border-box;
      }

      .container {
        width: 100%;
        background: #fff;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        padding: 20px;
        box-sizing: border-box;
        margin-bottom: 20px;
        position: relative;
        page-break-inside: avoid; /* Prevent page breaks inside the container */
      }

      /* Watermark Style */
      .watermark {
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%) rotate(-45deg); /* Rotate 45 degrees */
        font-size: 48px;
        font-weight: bold;
        color: rgba(0, 0, 0, 0.1); /* Default Light Gray */
        pointer-events: none;
        user-select: none;
        z-index: 0; /* Ensure watermark appears below content */
        white-space: nowrap;
      }

      .watermark.department {
        color: rgba(0, 0, 255, 0.1); /* Light Blue for Department's Copy */
      }

      .watermark.pao {
        color: rgba(255, 0, 0, 0.1); /* Light Red for Pao's Copy */
      }

      .header {
        text-align: center;
        margin-bottom: 20px;
        z-index: 1;
        position: relative;
      }

      .header h1 {
        margin: 0;
        font-size: 28px;
        color: #333;
      }

      .header h2 {
        margin: 0;
        font-size: 20px;
        font-weight: normal;
        color: #666;
      }

      .form-group {
        display: flex;
        flex-direction: row;
        margin-bottom: 15px;
        align-items: center;
        z-index: 1;
        position: relative;
      }

      .form-group label {
        min-width: 150px; /* Consistent left alignment */
        margin-right: 10px;
        color: #555;
        font-weight: bold;
        text-align: left;
      }

      .form-group input,
      .form-group textarea {
        width: 100%;
        max-width: 1200px;
        border: none;
        border-bottom: 1px solid #333;
        outline: none;
        padding: 5px;
        font-size: 14px;
        color: #333;
        background: transparent;
        resize: none;
      }

      .form-group input.no-border,
      .form-group textarea.no-border {
        border-bottom: none;
      }

      .form-group input.short-line {
        width: 75px; /* Half inch for Form No. */
      }

      .form-group input.medium-line {
        width: 125px; /* About 1 inch for others */
      }

      .form-group input.short-input {
        width: 200px; /* Shortened input width */
      }
      .form-group input.medium-input {
        width: 400px; /* Shortened input width */
      }

      .form-group input:focus,
      .form-group textarea:focus {
        border-bottom-color: #007BFF;
      }

      table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 40px; /* Increased margin to provide more space */
      }

      table, th, td {
        border: 1px solid #333;
      }

      th, td {
        padding: 10px;
        text-align: left;
        font-size: 14px;
      }

      .nature-of-work {
        word-wrap: break-word; /* Forces text to break within the cell */
        white-space: normal; /* Allows text to wrap to the next line */
        overflow-wrap: break-word; /* Handles breaking at appropriate points */
        max-width: 100%; /* Ensures text stays within cell bounds */
        display: block; /* Ensures block display for proper wrapping */
        /*width: 50%; /* Fixed width for Nature of Work column */
      }

      .list-of-materials {
        width: 50%; /* Fixed width for List of Material(s) column */
      }

      .signature-group {
        display: flex;
        flex-wrap: wrap;
        justify-content: space-between;
        margin-top: 50px; /* Increased margin to provide more space */
        margin-bottom: 10px;
        z-index: 1;
        position: relative;
      }

      .signature-group div {
        flex: 1 1 45%; /* Adjusts to 45% width to fit two columns */
        text-align: center;
        position: relative; /* Allows absolute positioning within each signature box */
        margin-bottom: 20px;
      }

      .signature-group div .line {
        border-bottom: 1px solid #333; /* Reduced border thickness */
        width: 80%;
        margin: 0 auto;
        padding: 2px 0; /* Reduced padding */
      }

      .signature-group div label {
        display: block;
        margin-top: 2px; /* Reduced margin */
        color: #555;
        font-size: 12px; /* Reduced font size */
      }

      .signature-group div .name {
        margin-bottom: 2mm; /* 2mm space below the name */
        font-size: 12px; /* Reduced font size */
        color: #333;
      }

      .print-button-container {
        text-align: center;
        margin-top: 20px;
      }

      .button-group {
        display: flex;
        justify-content: center;
        gap: 10px;
      }

      button {
        padding: 10px 20px;
        background-color: #007BFF;
        color: white;
        border: none;
        font-size: 16px;
        cursor: pointer;
        border-radius: 5px;
        transition: background-color 0.3s ease;
      }

      button:hover {
        background-color: #0056b3;
      }

      .footer {
        margin-top: 30px;
        text-align: center;
        color: #666;
        font-size: 14px;
      }

      .footer a {
        color: #007BFF;
        text-decoration: none;
        transition: color 0.3s ease;
      }

      .footer a:hover {
        color: #0056b3;
      }

      @media print {
        button {
          display: none;
        }

        body {
          padding: 0.25in; /* Adjusted margins to minimum */
        }

        .wrapper {
          width: 100%;
          max-width: none;
          padding: 0;
        }

        .container {
          width: 100%;
          max-width: none;
          box-shadow: none;
          border-radius: 0;
          padding: 0;
          margin: 0;
        }

        .header h1 {
          font-size: 20px;
        }

        .header h2 {
          font-size: 16px;
        }

        .form-group {
          margin-bottom: 10px;
        }

        .signature-group {
          margin-top: 50px; /* Increased margin to provide more space */
        }

        .footer {
          font-size: 12px;
        }

        .form-group input,
        .form-group textarea {
          font-size: 12px;
        }

        .signature-group div .line {
          font-size: 12px;
        }
      }

      .no-page-break {
        page-break-inside: avoid;
      }
    </style>
  </head>
  <body>
    <div class="wrapper">
      <div class="no-page-break">
        <!-- Content for departments copy and PAO's copy -->
        <div id="departments-copy">
          <!-- First Copy -->
          <div class="container">
            <div class="watermark department">Department's Copy</div>
            <div class="header">
              <h1>UNIVERSITY OF SAN JOSE-RECOLETOS</h1>
              <h2>Maintenance Job Order</h2>
            </div>
            <div class="form-group">
              <label for="formNumber1">Form No.:</label>
              <input type="text" id="formNumber1" class="short-input" value="${formData.formNumber}" readonly>
            </div>
            <div class="form-group">
              <label for="department1">Department:</label>
              <input type="text" id="department1" class="short-input" value="${formData.department}" readonly>
            </div>
            <div class="form-group">
              <label for="dateFiled1">Date/Time Requested:</label>
              <input type="text" id="dateFiled1" class="short-input" value="${formData.dateFiled}" readonly>
            </div>
            <div class="form-group">
              <label for="location1">Location:</label>
              <input type="text" id="location1" class="medium-input" value="${formData.location}" readonly>
            </div>
            <table>
              <tr>
                <th>Nature of Work</th>
                <th>List of Material(s)</th>
              </tr>
              <tr>
                <td class="nature-of-work">
                  <div>${formData.natureOfWork}</div>
                </td>
                <td class="list-of-materials"></td>
              </tr>
            </table>
            <div class="signature-group">
              <div>
                <div class="name">${formData.requestedBy}</div>
                <div class="line"></div>
                <label for="requestedBy1">Requested by</label>
              </div>
              <div>
                <div class="line"></div>
                <label for="notedBy1">Noted by: Campus Planning Head</label>
              </div>
              <div>
                <div class="name">Rev. Fr. Leander V. Barrot, OAR</div>
                <div class="line"></div>
                <label for="approvedBy1">Approved by</label>
              </div>
              <div>
                <div class="line"></div>
                <label for="dateCompleted1">Date/Time Completed</label>
              </div>
            </div>
          </div>
        </div>
        <div id="paos-copy">
          <!-- Second Copy -->
          <div class="container">
            <div class="watermark pao">PAO's Copy</div>
            <div class="header">
              <h1>UNIVERSITY OF SAN JOSE-RECOLETOS</h1>
              <h2>Maintenance Job Order</h2>
            </div>
            <div class="form-group">
              <label for="formNumber2">Form No.:</label>
              <input type="text" id="formNumber2" class="short-input" value="${formData.formNumber}" readonly>
            </div>
            <div class="form-group">
              <label for="department2">Department:</label>
              <input type="text" id="department2" class="short-input" value="${formData.department}" readonly>
            </div>
            <div class="form-group">
              <label for="dateFiled2">Date/Time Requested:</label>
              <input type="text" id="dateFiled2" class="short-input" value="${formData.dateFiled}" readonly>
            </div>
            <div class="form-group">
              <label for="location2">Location:</label>
              <input type="text" id="location2" class="medium-input" value="${formData.location}" readonly>
            </div>
            <table>
              <tr>
                <th>Nature of Work</th>
                <th>List of Material(s)</th>
              </tr>
              <tr>
                <td class="nature-of-work">
                  <div>${formData.natureOfWork}</div>
                </td>
                <td class="list-of-materials"></td>
              </tr>
            </table>
            <div class="signature-group">
              <div>
                <div class="name">${formData.requestedBy}</div>
                <div class="line"></div>
                <label for="requestedBy2">Requested by</label>
              </div>
              <div>
                <div class="line"></div>
                <label for="notedBy2">Noted by: Campus Planning Head</label>
              </div>
              <div>
                <div class="name">Rev. Fr. Leander V. Barrot, OAR</div>
                <div class="line"></div>
                <label for="approvedBy2">Approved by</label>
              </div>
              <div>
                <div class="line"></div>
                <label for="dateCompleted2">Date/Time Completed</label>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div class="print-button-container">
      <div class="button-group">
        <button onclick="window.print()">Print Here</button>
        <button id="downloadButton">Download as PDF</button>
      </div>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>

    <script>
      const formData = {
        formNumber: "${formData.formNumber}",
        department: "${formData.department}",
        dateFiled: "${formData.dateFiled}",
        location: "${formData.location}",
        natureOfWork: "${formData.natureOfWork}",
        requestedBy: "${formData.requestedBy}"
      };

      window.onload = function() {
        document.getElementById('formNumber1').value = formData.formNumber;
        document.getElementById('department1').value = formData.department;
        document.getElementById('dateFiled1').value = formData.dateFiled;
        document.getElementById('location1').value = formData.location;
        document.getElementById('formNumber2').value = formData.formNumber;
        document.getElementById('department2').value = formData.department;
        document.getElementById('dateFiled2').value = formData.dateFiled;
        document.getElementById('location2').value = formData.location;
      }
      
      document.getElementById("downloadButton").addEventListener("click", function () {
        const element = document.querySelector(".wrapper"); // Select the main container

        // Dynamically adjust the wrapper size to fit within the page
        const adjustForPDF = () => {
          // Slight scaling down to fit both copies on one page
          element.style.transform = "scale(0.85)"; // 85% scale for content fitting
          element.style.transformOrigin = "top left"; // Align to top left corner
          element.style.width = "8.5in"; // Set width to match letter-size page width
          element.style.height = "11in"; // Set height to match letter-size page height
          element.style.boxSizing = "border-box"; // Ensure no overflow from margins/padding
          element.style.marginTop = "0"; // Remove any top margin
        };

        adjustForPDF(); // Apply adjustments

        // PDF options
        const opt = {
          margin: [0.25, 0.25, 0.25, 0.25], // Top, right, bottom, left margins in inches
          filename: "Maintenance_Job_Order.pdf", // Name of the PDF file
          image: { type: "jpeg", quality: 0.98 },
          html2canvas: { scale: 2 }, // Higher scale for better resolution
          jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
          pagebreak: { mode: ["avoid-all"] }, // Avoid breaking content across pages
        };

        // Generate and download PDF
        html2pdf().set(opt).from(element).save().then(() => {
          // Reset styles after PDF generation to avoid affecting the layout on the page
          element.style.transform = "";
          element.style.width = "";
          element.style.height = "";
          element.style.boxSizing = "";
          element.style.marginTop = "";
        });
}

// ====== ANNOUNCEMENT FUNCTIONALITY ======
      });
    </script>
  </body>
</html>
  `;
}

// ====== ANNOUNCEMENT FUNCTIONALITY ======

// Get announcement data
function getAnnouncement() {
  try {
    Logger.log('getAnnouncement function called');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(ANNOUNCEMENTS_SHEET_NAME);
    
    Logger.log('Looking for sheet: ' + ANNOUNCEMENTS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Announcements sheet not found, creating it...');
      // Create the Announcements sheet if it doesn't exist
      sheet = ss.insertSheet(ANNOUNCEMENTS_SHEET_NAME);
      sheet.appendRow(['Title', 'Body', 'Date Created', 'Active']);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
      headerRange.setBorder(true, true, true, true, true, true);
      
      // Set column widths
      sheet.setColumnWidth(1, 200); // Title
      sheet.setColumnWidth(2, 400); // Body
      sheet.setColumnWidth(3, 150); // Date Created
      sheet.setColumnWidth(4, 100); // Active
      
      // Add sample announcement
      const timestamp = new Date();
      sheet.appendRow([
        'Welcome to PAO Basak Venue Management System',
        'This is the centralized system for managing venue reservations, job orders, and aircon concerns. Please follow the guidelines and submit your requests properly.',
        timestamp,
        'TRUE'
      ]);
      
      Logger.log('Sample announcement added to new sheet');
      
      // Return the sample announcement
      return {
        success: true,
        announcement: {
          title: 'Welcome to PAO Basak Venue Management System',
          body: 'This is the centralized system for managing venue reservations, job orders, and aircon concerns. Please follow the guidelines and submit your requests properly.',
          dateCreated: timestamp,
          active: true
        }
      };
    }

    Logger.log('Announcements sheet found, reading data...');
    const data = sheet.getDataRange().getValues();
    Logger.log('Sheet has ' + data.length + ' rows');
    
    // Look for the first active announcement (skip header row)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const title = row[0];
      const body = row[1];
      const dateCreated = row[2];
      const active = row[3];
      
      Logger.log('Row ' + i + ': title=' + title + ', body=' + body + ', active=' + active);
      
      // Check if announcement is active and has content
      if (active && (active.toString().toLowerCase() === 'true' || active === true) && 
          title && title.toString().trim() && body && body.toString().trim()) {
        Logger.log('Found active announcement: ' + title);
        return {
          success: true,
          announcement: {
            title: title.toString().trim(),
            body: body.toString().trim(),
            dateCreated: dateCreated,
            active: true
          }
        };
      }
    }
    
    Logger.log('No active announcement found');
    // No active announcement found
    return {
      success: true,
      announcement: null
    };
    
  } catch (error) {
    Logger.log('Error in getAnnouncement: ' + error.message);
    return { success: false, error: error.toString() };
  }
}