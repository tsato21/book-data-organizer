function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Create a custom menu in the Google Sheets UI
  ui.createMenu('Custom Menu')
    .addItem('確認メール_Mail Mergeデータ抽出', 'outputItemsForConfirmEmail')
    .addSeparator()
    .addItem('E-Book-Link共有メール_Mail Mergeデータ抽出', 'outputItemsForLinkShareEmail')
    .addToUi();
}

function outputItemsForConfirmEmail() {
   let ss = SpreadsheetApp.getActiveSpreadsheet();
   let sheet = ss.getSheetByName(REFERSHEET_NAME_CONFIRM_EMAIL); // Replace with your actual input sheet name

  // Find the last row with content in column A
  let lastRow = sheet.getRange("A2:A").getValues().filter(String).length;

  // Get the data range up to the last row with content in column A
  let data = sheet.getRange(2, 1, lastRow, sheet.getLastColumn()).getValues();

   let bookDataByFaculty = {};
   let outputSheet = ss.getSheetByName(OUTPUTSHEET_NAME_CONFIRM_EMAIL);
   let previousDataNum = outputSheet.getLastRow()-1;
   if(previousDataNum > 1){
    outputSheet.getRange(2,1,previousDataNum,3).clearContent();
   }

   // Start from row 2 to skip headers (assuming headers are in row 1)
   for (let i = 2; i < data.length; i++) {
       let row = data[i];
       let firstFacultyName = arrangeName_(row[0]);
       let secondFacultyName = arrangeName_(row[2]);
       let facultyName = (firstFacultyName) + "-sensei" + (secondFacultyName ? " and " + secondFacultyName + "-sensei" : "");
       let email = row[1] + (row[3] ? ", " + row[3] : "");
       let courseInfo = `【${row[4]} (${row[5]})】`;
       let bookDetails = "・Book Title: " + row[6] 
                        + "\n・Format: " + row[8] 
                        + "\n・ISBN: " + row[9] 
                        + "\n・eStore Link: " + row[10]
                        + "\n・NOTE: " + (row[7] || 'none');

       if (!bookDataByFaculty[facultyName]) {
           bookDataByFaculty[facultyName] = { email: email, allBookDetails: {} };
       }
       if (!bookDataByFaculty[facultyName].allBookDetails[courseInfo]) {
           bookDataByFaculty[facultyName].allBookDetails[courseInfo] = [];
       }
       bookDataByFaculty[facultyName].allBookDetails[courseInfo].push(bookDetails);
   }

   // Process to finalize details
   let outputData = [];
   for (let faculty in bookDataByFaculty) {
       let finalizedDetails = "";
       let allBookDetails = bookDataByFaculty[faculty].allBookDetails;
       let isFirstCourse = true;  // Flag to check if it's the first course

       for (let eachCourseBooks in allBookDetails) {
           let courseDetail = "";
           if (allBookDetails[eachCourseBooks].length > 1) {
               // Multiple book details
               let combinedDetails = "";
               allBookDetails[eachCourseBooks].forEach(function(detail, index) {
                   combinedDetails += "Book " + (index + 1) + ") \n" + detail + "\n";
               });
               courseDetail = eachCourseBooks + "\n" + combinedDetails;
           } else {
               // Single book detail
               courseDetail = eachCourseBooks + "\n" + allBookDetails[eachCourseBooks][0];
           }

           // Append course detail, avoid adding newline at the beginning
           finalizedDetails += (isFirstCourse ? "" : "\n\n") + courseDetail;
           isFirstCourse = false; // Update the flag after the first course
       }

       // Trim any extra newline characters at the end
       finalizedDetails = finalizedDetails.trim();

       outputData.push([faculty, bookDataByFaculty[faculty].email, finalizedDetails]);
   }

   // Output to a new sheet
   outputSheet.getRange(2, 1, outputData.length, 3)
                                .setValues(outputData)
                                .setVerticalAlignment("middle")
                                .setWrap(true);
}

function arrangeName_(name) {
   let parts = name.trim().split(" "); // Split the name into parts
   if (parts.length > 0) {
       let lastName = parts[parts.length - 1];
       return lastName.charAt(0).toUpperCase() + lastName.slice(1).toLowerCase();
   } else {
       return ""; // Return an empty string if the name is empty or invalid
   }
}


function outputItemsForLinkShareEmail() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REFERSHEET_NAME_LINKSHARE_EMAIL); // Replace with your actual input sheet name

  // Find the last row with content in column A
  let lastRow = sheet.getRange("A2:A").getValues().filter(String).length;

  const disCountCode = Browser.inputBox('Input the Discount Code provided by Cengage',Browser.Buttons.OK_CANCEL);
  const availableDate = Browser.inputBox('Input the available date provided by Cengage (e.g. 2023/9/30)',Browser.Buttons.OK_CANCEL);
  const datePattern = /^(20\d{2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12]\d|3[01])$/;

  if (disCountCode === 'cancel'){
    Browser.msgBox('Inputting the discount code is cancelled. Try again.');
    return;
  } else if (disCountCode === ''){
    Browser.msgBox('Discount code is empty. Try again.');
    return;
  } else if (availableDate === 'cancel'){
    Browser.msgBox('Inputting the available date is cancelled. Try again.');
    return;
  } else if (availableDate === ''){
    Browser.msgBox('Available Date is empty. Try again.');
    return;
  } else if (!datePattern.test(availableDate)) {
    Browser.msgBox('Available Date is invalid date format. Please enter a date in the format YYYY/MM/DD or YYYY/M/D.');
    return;
  }


  // Get the data range up to the last row with content in column A
  let data = sheet.getRange(2, 1, lastRow, sheet.getLastColumn()).getValues();

  let bookDataByFaculty = {};
  let outputSheet = ss.getSheetByName(OUTPUTSHEET_NAME_LINKSHARE_EMAIL);
  let previousDataNum = outputSheet.getLastRow() - 1;
  if (previousDataNum > 1) {
    outputSheet.getRange(2, 1, previousDataNum, 5).clearContent();
  }

  for (let i = 2; i < data.length; i++) {
    let row = data[i];
    let facultyName = arrangeName_(row[2]) + "-sensei";
    let email = row[0];
    let courseInfo = `【${row[3]} (${row[4]})】`;
    let bookDetails = {
      title: row[12],
      price: Number(row[16]).toLocaleString() + " yen",
      link: row[17]
    };

    if (!bookDataByFaculty[facultyName]) {
      bookDataByFaculty[facultyName] = { email: email, allBookDetails: {} };
    }
    if (!bookDataByFaculty[facultyName].allBookDetails[courseInfo]) {
      bookDataByFaculty[facultyName].allBookDetails[courseInfo] = [bookDetails];
    } else {
      // Add to existing book titles for the same courseInfo
      bookDataByFaculty[facultyName].allBookDetails[courseInfo].push(bookDetails);
    }
  }

  let outputData = [];
  for (let faculty in bookDataByFaculty) {
    let finalizedDetails = "";
    let allBookDetails = bookDataByFaculty[faculty].allBookDetails;
    let isFirstCourse = true;

    for (let eachCourseBooks in allBookDetails) {
      let courseDetail = eachCourseBooks + "\n";
      let bookTitles = [];
      let commonPrice = "";
      let commonLink = "";

      allBookDetails[eachCourseBooks].forEach(function (detail, index) {
        bookTitles.push(detail.title);
        if (index === 0) {
          commonPrice = "・Estore Price (After discount with tax): " + detail.price;
          commonLink = "\n・Estore Link: " + detail.link;
        }
      });

      courseDetail += "・Book Title(s): " + bookTitles.join(" / ") + "\n" + commonPrice + commonLink;
      finalizedDetails += (isFirstCourse ? "" : "\n\n") + courseDetail;
      isFirstCourse = false;
    }

    finalizedDetails = finalizedDetails.trim();
    outputData.push([faculty, bookDataByFaculty[faculty].email, finalizedDetails,disCountCode, availableDate]);
  }

  outputSheet.getRange(2, 1, outputData.length, 5)
    .setValues(outputData)
    .setVerticalAlignment("middle")
    .setWrap(true);
}

