This project involves a Google Apps Script that automates the process of organizing e-book data in Google Sheets and utilizing Mail Merge to send personalized emails. The script is specifically designed to handle two types of emails: confirmation emails and link share emails.

- Confirmation Emails: The script processes data to generate unique items for each confirmation email, such as recipient names, course information, and book details. It then uses mail merge to send personalized confirmation emails to the recipients.

- Link Share Emails: The script also handles data for link share emails. It organizes the necessary information, including recipient names, course details, and shared links. Using mail merge, it sends personalized link share emails to the recipients, providing them with the relevant links to access e-books or resources.

By automating these processes, the Google Apps Script streamlines the management and distribution of e-book data, making it easier to send personalized emails to recipients in bulk.

## Prerequisites

- A Google account with access to Google Sheets.
- Basic understanding of Google Sheets and Google Apps Script.

## Setup

1. **Open Your Google Sheet**: Access [a sample Google Sheet](https://docs.google.com/spreadsheets/d/1mMuQSK06hIcAUcI1qW4cgD2_IKOXU9_DAR2CTj2a-a8/edit#gid=1834592607) where you want to use this script.
2. **Conduct GAS Authorizatio**: Access `Initial Setting` Sheet and click the initial setting button. This enables you to go to the authorization page for Google Apps Script.
  ![Image of Initial Setting](assets/initial-setting.png)

## Usage

1. **Input Reference Data**: Make sure to input reference data in the orange range which is designated for each column in reference sheets (`Confirm Mail-Ref Data`/ `Link Share Mail-Ref Data`).
  ![Image of Input Reference Data](assets/input-ref-data.png)

2. **Process data**: Execute functions from custom menu in your Google Sheets UI.
  ![Image of Process Data](assets/output-data.png)


  - Select `Output Mail Merge Data for Confirm Email` to process data for confirmation emails.

    This function is responsible for processing data to generate mail merge data specifically for confirmation emails. The logic of the function can be summarized as follows:

    1. Read each row of data from the input source.
    2. If the second faculty info is provided, combine the names and emails of the faculty members.
    3. If the first faculty has multiple records, combine the data to create a single entry for that faculty.
    4. If a faculty has two books for a certain course, list the book information as "BookXX" where XX represents the book number.
    5. Generate the final mail merge data that can be used to send confirmation emails.

    Below are example screenshots illustrating the input data and the output data after applying the logic:

    ![Image of Input Data for Confirm](assets/ref-data-confirm.png)
    ![Image of Output Data for Confirm](assets/output-data-confirm.png)

  - Select `Output Mail Merge Data for Link Share Email` to process data for link share emails.

    This function handles the processing of data to generate mail merge data specifically for link share emails. The logic of the function can be summarized as follows:

    1. Read each row of data from the input source. (Unlike the confirmation email, this function does not accept the second faculty info.)
    2. If the faculty has multiple records, combine the data to create a single entry for that faculty.
    3. If a faculty has two books for a certain course, combine the the book information as "Book A/ Book B".
    4. Generate the final mail merge data that can be used to send confirmation emails.

    Below are example screenshots illustrating the input data and the output data after applying the logic:

    ![Image of Input Data for E-link Share](assets/ref-data-elink.png)
    ![Image of Output Data for E-link Share](assets/output-data-elink.png)
3. **Use Output Data for Mail Merge**: Processed data are output in output sheets (`Confirm Mail-Mail Merge Data`/ `Link Share Mail-Mail Merge Data`). Use the data for Mail Merge. As for the usage of Mail Merge, click HERE.


## Others

- **Clear Sample Data**: Clear sample data by clicking `Clear Sample Data` here.
  ![Image of Clear Sample Data Image](assets/clear-sample-data.png)


- **Customization**: You can customize the scripts according to your specific preferences.
