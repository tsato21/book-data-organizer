/**
 * This function can be used to manually trigger the authorization flow for the script.
 */
function showAuthorizationDialog() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  Browser.msgBox('Authorization has been granted. You can now use the script functionalities.', Browser.Buttons.OK);
}

/**
 * Creates a custom menu in the Google Sheets UI when the spreadsheet is opened.
 * This function is automatically triggered when the spreadsheet containing the script is opened.
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Create a custom menu in the Google Sheets UI
  ui.createMenu('Custom Menu')
    .addItem('Output Mail Merge Data for Confirm Email', 'outputItemsForConfirmEmail')
    .addSeparator()
    .addItem('Output Mail Merge Data for Link Share Email', 'outputItemsForLinkShareEmail')
    .addToUi();
}

/**
 * Processes data from a specific sheet and outputs it for use in confirm email mail merge.
 * This function reads data from a predefined sheet, processes it, and outputs it in a format
 * suitable for mail merge operations.
 */
function outputItemsForConfirmEmail() {
  try{
      let ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(REFERSHEET_NAME_CONFIRM_EMAIL); // Replace with your actual input sheet name

      // Find the last row with content in column A
      let lastRow = sheet.getRange("A2:A").getValues().filter(String).length;

      // Get the data range up to the last row with content in column A
      let data = sheet.getRange(2, 1, lastRow, sheet.getLastColumn()).getValues();
      console.log(data);

      let bookDataByFaculty = {};
      let outputSheet = ss.getSheetByName(OUTPUTSHEET_NAME_CONFIRM_EMAIL);
      let previousDataNum = outputSheet.getLastRow()-1;
      if(previousDataNum > 1){
        outputSheet.getRange(2,1,previousDataNum,3).clearContent();
      }

      // Read each row data
      for (let i = 0; i < data.length; i++) {
          let row = data[i];
          //Get each cell data
          let firstFacultyName = arrangeName_(row[0]);
          let secondFacultyName = arrangeName_(row[2]);
          //Combie faculty names if the second faculty is input. If not, only the first faculty name
          let facultyName = (firstFacultyName) + "-sensei" + (secondFacultyName ? " and " + secondFacultyName + "-sensei" : "");
          //Combine emails if the second faculty is input. If not, only the email for the first faculty 
          let email = row[1] + (row[3] ? ", " + row[3] : "");
          let courseCode = row[4];
          let courseTitle = row[5];
          let courseInfo = `【${courseCode} (${courseTitle}】`;
          let bookTitle = row[6];
          let note = row[7];
          let digitalFormat = row[8];
          let isbn = row[9];
          let link = row[10];

          let bookDetails = "・Book Title: " + bookTitle 
                            + "\n・Format: " + digitalFormat
                            + "\n・ISBN: " + isbn
                            + "\n・eStore Link: " + link
                            + "\n・NOTE: " + (note || 'none');
          
          //If the bookDataByFaculty obect does not have the target faculty property, create a new object for that property and new object.
          if (!bookDataByFaculty[facultyName]) {
              bookDataByFaculty[facultyName] = { email: email, allBookDetails: {} };
          }
          //If allBookDetails by the particular faculty do not have the courseInfo properry, make a new array.
          if (!bookDataByFaculty[facultyName].allBookDetails[courseInfo]) {
              bookDataByFaculty[facultyName].allBookDetails[courseInfo] = [];
          }
          bookDataByFaculty[facultyName].allBookDetails[courseInfo].push(bookDetails);
      }

      // Process to finalize details since some of the data should be combined
      let outputData = [];
      //Refer to bookDataByFaculty object by each faculty property
      for (let faculty in bookDataByFaculty) {
          let finalizedDetails = "";
          let allBookDetails = bookDataByFaculty[faculty].allBookDetails;
          let isFirstCourse = true;  // Flag to check if it's the first course to put "\n" (new line) between courses

          for (let courseInfo in allBookDetails) {
            //courseInfo is a property that has bookDetails values
              let courseDetail = "";
              if (allBookDetails[courseInfo].length > 1) {
                  // Multiple bookDetails for certain course. In this case, put BookX) ahead of each courseInfo.
                  let combinedDetails = "";
                  allBookDetails[courseInfo].forEach(function(detail, index) {
                      combinedDetails += "Book " + (index + 1) + ") \n" + detail + "\n";
                  });
                  courseDetail = courseInfo + "\n" + combinedDetails;
              } else {
                  // Single book detail. In this case, simply put courseInfo and bookDetails (first array)
                  courseDetail = courseInfo + "\n" + allBookDetails[courseInfo][0];
              }

              // Append course detail, avoid adding newline at the beginning
              finalizedDetails += (isFirstCourse ? "" : "\n\n") + courseDetail;
              isFirstCourse = false; // Update the flag after the first course
          }

          // Trim any extra newline characters at the end
          finalizedDetails = finalizedDetails.trim();
          
          outputData.push([faculty, bookDataByFaculty[faculty].email, finalizedDetails]);
      }

      // Output the data into the designated sheet with some format organized
      outputSheet.getRange(2, 1, outputData.length, 3)
                                    .setValues(outputData)
                                    .setVerticalAlignment("middle")
                                    .setWrap(true);

      Browser.msgBox(`Confirmation email data for mail merge has been successfully output on the sheet, "${OUTPUTSHEET_NAME_CONFIRM_EMAIL}".`, Browser.Buttons.OK);
      return;
  } catch(e){
      console.error(`Error while processing ${REFERSHEET_NAME_CONFIRM_EMAIL}: ${e.toString()} at ${e.stack}`);
      Browser.msgBox(`Error while processing ${REFERSHEET_NAME_CONFIRM_EMAIL}. Contact the owner of this script.`);
  }
}

/**
 * Arranges the provided name in a specific format.
 * Extracts and capitalizes the last name from a full name string.
 * 
 * @param {string} name - The full name to be arranged.
 * @returns {string} The arranged name.
 * 
 * @example
 * // returns "Doe"
 * arrangeName_("John Doe");
 * 
 * @example
 * // returns "Smith"
 * arrangeName_("Jane Smith");
 */
function arrangeName_(name) {
   let parts = name.trim().split(" "); // Split the name into parts
   if (parts.length > 0) {
       let lastName = parts[parts.length - 1];
       return lastName.charAt(0).toUpperCase() + lastName.slice(1).toLowerCase();
   } else {
       return ""; // Return an empty string if the name is empty or invalid
   }
}


/**
 * Processes data from a specific sheet and outputs it for use in link share email mail merge.
 * This function reads data from a predefined sheet, collects additional information through
 * user input, processes it, and outputs it in a format suitable for mail merge operations.
 */
function outputItemsForLinkShareEmail() {
  try{
      let ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(REFERSHEET_NAME_LINKSHARE_EMAIL); // Replace with your actual input sheet name

      // Find the last row with content in column A
      let lastRow = sheet.getRange("C2:C").getValues().filter(String).length;
      // console.log(lastRow);

      const disCountCode = Browser.inputBox(`Input the Discount Code provided by ${COMPANY_NAME}`,Browser.Buttons.OK_CANCEL);
      const availableDate = Browser.inputBox(`Input an expiry date provided by ${COMPANY_NAME} (e.g. 2023/9/30)`,Browser.Buttons.OK_CANCEL);
      const datePattern = /^(20\d{2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12]\d|3[01])$/;

      //If the input data is cancelled or invalid, return and ask the user to try again.
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

      for (let i = 0; i < data.length; i++) {
        let row = data[i];
        let facultyName = arrangeName_(row[0]) + "-sensei";
        let email = row[1];
        let courseCode = row[2];
        let courseTitle = row[3];
        let courseInfo = `【${courseCode} (${courseTitle})】`;
        let bookTitile = row[4];
        let price = Number(row[7]).toLocaleString() + " yen";
        let link = row[8];
        let bookDetails = {
          title: bookTitile,
          price: price,
          link: link
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

        for (let courseInfo in allBookDetails) {
          let courseDetail = courseInfo + "\n";
          let bookTitles = [];
          let commonPrice = "";
          let commonLink = "";

          allBookDetails[courseInfo].forEach(function (detail, index) {
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

      Browser.msgBox(`Link share data for mail merge has been successfully output on the sheet, "${OUTPUTSHEET_NAME_LINKSHARE_EMAIL}".`, Browser.Buttons.OK);
      return;
  } catch(e){
      console.error(`Error while processing ${REFERSHEET_NAME_LINKSHARE_EMAIL}: ${e.toString()} at ${e.stack}`);
      Browser.msgBox(`Error while processing ${REFERSHEET_NAME_LINKSHARE_EMAIL}. Contact the owner of this script.`);
  }

}