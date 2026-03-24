/**
 * Email Functions - UPDATED for new quote structure
 */

function sendQuoteEmail(data) {
  try {
    var recipient = data.email;
    
    if (!recipient) {
      return {success: false, message: 'No email address provided'};
    }
    
    var type = data.type || 'quote';
    var subject = type === 'invoice' ?
      'Invoice #' + data.number + ' from Your Company' :
      'Quote #' + data.number + ' from Your Company';
    
    var htmlBody = createEmailBody(data);
    
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });
    
    return {success: true, message: 'Email sent successfully to ' + recipient};
  } catch (error) {
    return {success: false, message: 'Error sending email: ' + error.toString()};
  }
}

// Create HTML email body - UPDATED for new structure
function createEmailBody(data) {
  var type = data.type || 'quote';
  var title = type === 'invoice' ? 'INVOICE' : 'QUOTE';
  var number = data.number;
  
  // Use new items structure or fall back to old lineItems
  var items = data.items || data.lineItems || [];
  
  var html = '<html><body style="font-family: Arial, sans-serif; color: #333;">';
  html += '<div style="max-width: 800px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px;">';
  html += '<div style="text-align: center; background: #001f3f; color: white; padding: 20px; border-radius: 8px 8px 0 0;">';
  html += '<h1 style="margin: 0;">' + title + '</h1>';
  html += '<h2 style="margin: 10px 0 0 0;">#' + number + '</h2>';
  html += '</div>';
  
  html += '<div style="padding: 20px;">';
  html += '<p>Dear ' + data.customerName + ',</p>';
  html += '<p>Please find your ' + type + ' details below:</p>';
  
  // Show objective if present
  if (data.objective) {
    html += '<div style="margin: 20px 0; padding: 15px; background: #f8f9fa; border-left: 4px solid #001f3f;">';
    html += '<h3 style="margin: 0 0 10px 0;">Project Objective</h3>';
    html += '<p style="margin: 0;">' + data.objective + '</p>';
    html += '</div>';
  }
  
  html += '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">';
  html += '<thead>';
  html += '<tr style="background: #001f3f; color: white;">';
  html += '<th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Item</th>';
  html += '<th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Model Number</th>';
  html += '<th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Item Name</th>';
  html += '<th style="padding: 12px; text-align: center; border: 1px solid #ddd;">QTY</th>';
  html += '</tr>';
  html += '</thead>';
  html += '<tbody>';
  
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    html += '<tr style="border-bottom: 1px solid #ddd;">';
    html += '<td style="padding: 10px; border: 1px solid #ddd;">' + (i + 1) + '</td>';
    html += '<td style="padding: 10px; border: 1px solid #ddd;">' + (item.modelNumber || '') + '</td>';
    html += '<td style="padding: 10px; border: 1px solid #ddd;">' + (item.itemName || item.description || '') + '</td>';
    html += '<td style="padding: 10px; text-align: center; border: 1px solid #ddd;">' + (item.qty || 0) + '</td>';
    html += '</tr>';
  }
  
  html += '</tbody>';
  html += '</table>';
  
  html += '<div style="text-align: right; padding: 20px; background: #f5f5f5; border-radius: 5px; margin-top: 20px;">';
  
  // Use new structure if available
  if (data.materialCost !== undefined) {
    html += '<div style="margin: 5px 0;"><strong>Material Cost:</strong> $' + parseFloat(data.materialCost || 0).toFixed(2) + '</div>';
    html += '<div style="margin: 5px 0;"><strong>Installation & Training:</strong> $' + parseFloat(data.installationCost || 0).toFixed(2) + '</div>';
    html += '<div style="margin: 10px 0; padding-top: 10px; border-top: 1px solid #ddd;"><strong>Sub Total:</strong> $' + parseFloat(data.subTotal || 0).toFixed(2) + '</div>';
    html += '<div style="margin: 5px 0;"><strong>Sales Tax:</strong> $' + parseFloat(data.salesTax || 0).toFixed(2) + '</div>';
    html += '<div style="margin: 10px 0 0 0; font-size: 20px; color: #001f3f;"><strong>GRAND TOTAL: $' + parseFloat(data.grandTotal || 0).toFixed(2) + '</strong></div>';
    
    html += '<div style="margin-top: 20px; padding-top: 15px; border-top: 2px solid #001f3f;">';
    html += '<h4 style="margin: 0 0 10px 0;">Payment Schedule</h4>';
    html += '<div style="margin: 5px 0;"><strong>Down Payment:</strong> $' + parseFloat(data.downPayment || 0).toFixed(2) + '</div>';
    html += '<div style="margin: 5px 0;"><strong>Final Payment:</strong> $' + parseFloat(data.finalPayment || 0).toFixed(2) + '</div>';
    html += '</div>';
  } else {
    // Old structure fallback
    html += '<div style="margin: 5px 0;"><strong>Subtotal:</strong> $' + parseFloat(data.subtotal || 0).toFixed(2) + '</div>';
    html += '<div style="margin: 5px 0;"><strong>Tax:</strong> $' + parseFloat(data.taxAmount || 0).toFixed(2) + '</div>';
    if (data.other && parseFloat(data.other) > 0) {
      html += '<div style="margin: 5px 0;"><strong>Other:</strong> $' + parseFloat(data.other).toFixed(2) + '</div>';
    }
    html += '<div style="margin: 10px 0 0 0; font-size: 20px; color: #001f3f;"><strong>TOTAL: $' + parseFloat(data.total || 0).toFixed(2) + '</strong></div>';
  }
  
  html += '</div>';
  
  html += '<p style="margin-top: 30px;">Thank you for your business!</p>';
  html += '<p style="color: #666; font-size: 12px; margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd;">This is an automated email from Your Company</p>';
  html += '</div>';
  html += '</div>';
  html += '</body></html>';
  
  return html;
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}