function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var action = e.parameter.action;
  
  if (action === "summary") {
    var lastRow = sheet.getLastRow();
    // Exclude header row if the sheet has data
    var count = lastRow > 1 ? lastRow - 1 : 0;
    
    // Add offset if you want a baseline number of signups displayed
    var displayCount = count + 247; // Start count from 247 as baseline
    
    return ContentService.createTextOutput(JSON.stringify({
      "total_leads": displayCount
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput("Invalid Action");
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Create headers if the sheet is brand new and empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", 
        "Name", 
        "Phone Number", 
        "LGA", 
        "WhatsApp Group Link", 
        "Source", 
        "Medium", 
        "Campaign", 
        "Referrer", 
        "Full URL"
      ]);
    }
    
    var data = JSON.parse(e.postData.contents);
    
    // Append the lead data to the Google Sheet
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name,
      data.phone,
      data.lga,
      data.whatsapp_group,
      data.source,
      data.medium,
      data.campaign,
      data.referrer,
      data.full_url
    ]);
    
    // Send email notification
    var emailRecipient = data.notify_email || "womantowomannetwork21@gmail.com";
    var emailSubject = "🎉 New Signup: " + data.name + " (" + data.lga + ")";
    
    var emailBody = 
      "<h3>New Sister Joined the Network!</h3>" +
      "<p><strong>Name:</strong> " + data.name + "</p>" +
      "<p><strong>Phone:</strong> " + data.phone + "</p>" +
      "<p><strong>LGA:</strong> " + data.lga + "</p>" +
      "<p><strong>WhatsApp Group:</strong> <a href='" + data.whatsapp_group + "'>" + data.whatsapp_group + "</a></p>" +
      "<br>" +
      "<p><strong>Source (UTM):</strong> " + data.source + "</p>" +
      "<p><strong>Medium:</strong> " + data.medium + "</p>" +
      "<p><strong>Campaign:</strong> " + data.campaign + "</p>" +
      "<p><strong>Referrer:</strong> " + data.referrer + "</p>" +
      "<p><strong>Time:</strong> " + (data.timestamp ? new Date(data.timestamp).toLocaleString() : new Date().toLocaleString()) + "</p>";
      
    MailApp.sendEmail({
      to: emailRecipient,
      subject: emailSubject,
      htmlBody: emailBody
    });
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "message": "Lead saved and email sent!"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
