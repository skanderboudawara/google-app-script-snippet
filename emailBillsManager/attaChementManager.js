function attachement_management() {
  /*
  This function will search for all the emails from the list of email addresses

  :param: None

  :return: None

  */
  Object.keys(LIST_AND_FOLDERS).forEach(function (entry) {
    /*
    This forEach will search for all the emails from the list of email addresses

    :param: entry: email address

    :return: None
    
    */
    Logger.log(`Treating ${entry}`)
    var allEmailThreads = GmailApp.search('from:' + entry) // get all emails from the list of email addresses
    var parentFolder = DriveApp.getFolderById(LIST_AND_FOLDERS[entry]); // get the folder where to save the attachments


    for (var thread = 0; thread < allEmailThreads.length; thread++) {
      messages = allEmailThreads[thread].getMessages() // get all the messages from the email thread
      for (var message_index = 0; message_index < messages.length; message_index++) {
        try {
          message = messages[message_index]
          year_message = message.getDate().getFullYear()
          month_message = message.getDate().getMonth() + 1
          save_folder = checkFolder(parentFolder, year_message)
          sub_folder_save_name = month_message + ' - ' + year_message
          sub_save_folder = checkFolder(save_folder, sub_folder_save_name)
          var all_attachments = message.getAttachments(); // get all the attachments from the message
          if (!all_attachments){
            continue;
          }
          all_attachments.forEach(function (attachment_file) {
            fileName = month_message + '/' + year_message + ' - ' + attachment_file.getName(); // create the name of the file
            if (checkFile(sub_save_folder, fileName)) {
              return; // if the file already exists, do not save it
            }
            if (attachment_file.getName().includes('pdf')) { 
              sub_save_folder.createFile(attachment_file.copyBlob()).setName(fileName);  // save the file
              }
          });
        }
        catch (e) {
          Logger.log(e)
        }
      }

    }

  });

};

