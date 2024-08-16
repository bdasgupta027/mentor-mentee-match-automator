function matchMentorsToMentees() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    const menteeSheet = spreadsheet.getSheetByName('Mentee Schedules'); 
    const mentorSheet = spreadsheet.getSheetByName('Mentor Schedules'); 
    let assignmentsSheet = spreadsheet.getSheetByName('Assignments');
  
    if (!assignmentsSheet) {
      assignmentsSheet = spreadsheet.insertSheet('Assignments');
    }
  
    if (!menteeSheet || !mentorSheet) {
      Logger.log('Mentee or Mentor sheet not found');
      return;
    }
  
    const menteeData = menteeSheet.getDataRange().getValues();
    const mentorData = mentorSheet.getDataRange().getValues();
  
    const mentees = menteeData.slice(1).map(row => ({
      name: row[1],
      sessionTimes: parseMultiLineText(row[2]), 
      preferredMentor: row[3], 
      email: row[4],
      course: row[5],
      position: row[6]
    }));
  
    const mentors = mentorData.slice(1).map(row => ({
      name: row[1],
      availability: expandAvailability(parseMultiLineText(row[2])) 
    }));
  
    const incompatiblePairs = [
      'tutor15 + tutor23',
      'tutor17 + tutor22',
      'tutor12 + Mentor1 Lastname',
      'tutor6 + tutor8',
      'tutor7 + tutor19',
      'tutor10 + tutor12'
    ].map(pair => pair.split(' + ').map(name => name.trim()));
  
    // Generate the first set of final assignments
    generateFinalAssignments(spreadsheet, mentors, mentees, incompatiblePairs, 'Final Assignments');
  
    // Generate the second set of final assignments with randomness
    generateFinalAssignments(spreadsheet, mentors, mentees, incompatiblePairs, 'Final Assignments Option 2', true);
  }
  
  function generateFinalAssignments(spreadsheet, mentors, mentees, incompatiblePairs, sheetName, randomize = false) {
    const mentorMenteeMap = {};
  
    mentors.forEach(mentor => {
      mentorMenteeMap[mentor.name] = [];
    });
  
    let unassignedMentees = [...mentees];
  
    function countSILeaders(mentorName) {
      return mentorMenteeMap[mentorName].filter(mentee => mentee.position === 'SI Leader').length;
    }
  
    function isIncompatible(mentorName, menteeName) {
      const mentorGroup = mentorMenteeMap[mentorName];
      return mentorGroup.some(existingMentee =>
        incompatiblePairs.some(pair =>
          (pair.includes(existingMentee.name) && pair.includes(menteeName)) ||
          (pair[0] === menteeName && pair[1] === mentorName)
        )
      );
    }
  
    if (randomize) {
      unassignedMentees = shuffleArray(unassignedMentees);
    }
  
    unassignedMentees = unassignedMentees.filter(mentee => {
      let assigned = false;
      if (mentee.preferredMentor) {
        const preferredMentor = mentors.find(mentor => mentor.name === mentee.preferredMentor);
        if (preferredMentor && preferredMentor.availability.some(slot => mentee.sessionTimes.includes(slot))) {
          if (mentorMenteeMap[preferredMentor.name].length < 6 && countSILeaders(preferredMentor.name) < 2) {
            if (!isIncompatible(preferredMentor.name, mentee.name)) {
              mentorMenteeMap[preferredMentor.name].push(mentee);
              assigned = true;
            }
          }
        }
      }
      return !assigned;
    });
  
    unassignedMentees.forEach(mentee => {
      for (let mentor of mentors) {
        if (mentorMenteeMap[mentor.name].length < 6) {
          const availableSlot = mentor.availability.find(slot => mentee.sessionTimes.includes(slot));
          if (availableSlot) {
            if (mentee.position !== 'SI Leader' || countSILeaders(mentor.name) < 2) {
              if (!isIncompatible(mentor.name, mentee.name)) {
                mentorMenteeMap[mentor.name].push(mentee);
                unassignedMentees = unassignedMentees.filter(m => m.name !== mentee.name);
                break;
              }
            }
          }
        }
      }
    });
  
    // Any remaining unassigned mentees should be distributed among mentors with fewer than 6 mentees
    unassignedMentees.forEach(mentee => {
      for (let mentorName in mentorMenteeMap) {
        if (mentorMenteeMap[mentorName].length < 6) {
          if (!isIncompatible(mentorName, mentee.name)) {
            mentorMenteeMap[mentorName].push(mentee);
            break;
          }
        }
      }
    });
  
    createFinalAssignmentSheet(spreadsheet, sheetName, mentorMenteeMap);
  }
  
  function createFinalAssignmentSheet(spreadsheet, sheetName, mentorMenteeMap) {
    let finalSheet = spreadsheet.getSheetByName(sheetName);
  
    if (!finalSheet) {
      finalSheet = spreadsheet.insertSheet(sheetName);
    } else {
      finalSheet.clear();
    }
  
    finalSheet.appendRow(['Mentor Name', 'Mentee Name', 'Email', 'Course', 'Position']);
  
    const mentorAssignments = mentorMenteeMap;
  
    Object.keys(mentorAssignments).forEach(mentorName => {
      finalSheet.appendRow([mentorName, '', '', '', '']);
      mentorAssignments[mentorName].forEach(mentee => {
        finalSheet.appendRow(['', mentee.name, mentee.email, mentee.course, mentee.position]);
        const menteeCell = finalSheet.getRange(finalSheet.getLastRow(), 2);
        if (mentee.preferredMentor === mentorName) {
          menteeCell.setFontWeight('bold');
        }
      });
    });
  }
  
  function parseMultiLineText(multiLineText) {
    if (!multiLineText) return [];
    return multiLineText.split(',').map(timeSlot => timeSlot.trim());
  }
  
  function expandAvailability(availability) {
    const expanded = [];
    availability.forEach(range => {
      const [day, times] = range.split(' ');
      const [start, end] = times.split('-');
      const startTime = parseTime(start);
      const endTime = parseTime(end);
  
      for (let time = startTime; time + 1 <= endTime; time++) {
        expanded.push(`${day} ${formatTime(time)}-${formatTime(time + 1)}`);
      }
    });
    return expanded;
  }
  
  function parseTime(time) {
    const [hour, minute] = time.split(':');
    const isPM = minute.toLowerCase().includes('pm');
    const parsedHour = parseInt(hour) + (isPM && hour !== '12' ? 12 : 0) - (isPM && hour === '12' ? 12 : 0) + (!isPM && hour === '12' ? -12 : 0);
    return parsedHour + (minute.includes('30') ? 0.5 : 0);
  }
  
  function formatTime(time) {
    const hour = Math.floor(time);
    const minute = (time % 1) * 60;
    const period = hour >= 12 ? 'pm' : 'am';
    const formattedHour = hour % 12 === 0 ? 12 : hour % 12;
    const formattedMinute = minute === 0 ? '00' : '30';
    return `${formattedHour}:${formattedMinute}${period}`;
  }
  
  function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }
  