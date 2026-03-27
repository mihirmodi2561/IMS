// Generate PDF for quote
function generateQuotePDF(quoteNumber) {
  try {
    var result = getQuoteByNumber(quoteNumber);
    
    if (!result.success) {
      return result;
    }

    var quote = result.quote;
    var html = createPDFHTML(quote, 'quote');
    var blob = Utilities.newBlob(html, 'text/html', 'quote.html');
    var pdf = blob.getAs('application/pdf').setName('Quote_' + quoteNumber + '.pdf');

    return {
      success: true,
      pdf: Utilities.base64Encode(pdf.getBytes()),
      filename: 'Quote_' + quoteNumber + '.pdf'
    };
  } catch (error) {
    return {success: false, message: 'Error generating PDF: ' + error.toString()};
  }
}

// Generate PDF for invoice
function generateInvoicePDF(invoiceNumber) {
  try {
    var result = getInvoiceByNumber(invoiceNumber);
    
    if (!result.success) {
      return result;
    }

    var invoice = result.invoice;
    var html = createPDFHTML(invoice, 'invoice');
    var blob = Utilities.newBlob(html, 'text/html', 'invoice.html');
    var pdf = blob.getAs('application/pdf').setName('Invoice_' + invoiceNumber + '.pdf');

    return {
      success: true,
      pdf: Utilities.base64Encode(pdf.getBytes()),
      filename: 'Invoice_' + invoiceNumber + '.pdf'
    };
  } catch (error) {
    return {success: false, message: 'Error generating PDF: ' + error.toString()};
  }
}

// Create HTML for PDF - UPDATED for new structure
function createPDFHTML(data, type) {
  var isInvoice = type === 'invoice';
  var title = isInvoice ? 'INVOICE' : 'QUOTE';
  var number = isInvoice ? data.invoiceNumber : data.quoteNumber;
  var dateLabel = isInvoice ? 'DATE' : 'DATE';
  var dateValue = data.date ? new Date(data.date).toLocaleDateString() : 'N/A';
  var secondDateLabel = isInvoice ? 'DUE DATE' : 'VALID UNTIL';
  var secondDateValue = isInvoice ?
    (data.dueDate ? new Date(data.dueDate).toLocaleDateString() : 'N/A') :
    (data.validUntil ? new Date(data.validUntil).toLocaleDateString() : 'N/A');

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
  html += '<style>';
  html += 'body { font-family: Arial, sans-serif; margin: 20px; color: #333; }';
  html += '.header { display: flex; justify-content: space-between; margin-bottom: 30px; }';
  html += '.company-info { flex: 1; }';
  html += '.company-info h1 { color: #001f3f; margin: 0; font-size: 24px; }';
  html += '.company-info p { margin: 5px 0; font-size: 12px; }';
  html += '.quote-details { text-align: right; }';
  html += '.quote-details h2 { color: #001f3f; margin: 0 0 15px 0; font-size: 28px; }';
  html += '.header-table { border-collapse: collapse; }';
  html += '.header-table td { padding: 5px 10px; font-size: 12px; }';
  html += '.header-table td:first-child { font-weight: bold; text-align: right; }';
//  html += '.section-header { background: #001f3f; color: white; padding: 8px; font-weight: bold; margin: 20px 0 10px 0; }';
  html += '.section-header { color: #001f3f; font-weight: bold; margin: 20px 0 10px 0; }';
  html += '.customer-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px; }';
  html += '.customer-field label { font-weight: bold; font-size: 11px; color: #666; display: block; margin-bottom: 3px; }';
  html += '.customer-field div { font-size: 12px; padding: 5px; border-bottom: 1px solid #ddd; }';
  html += '.objective-box { background: #f8f9fa; padding: 15px; margin: 20px 0; border-left: 4px solid #001f3f; }';
  html += '.objective-box h3 { margin: 0 0 10px 0; font-size: 14px; color: #001f3f; }';
  html += '.objective-box p { margin: 0; font-size: 12px; line-height: 1.6; }';
  html += 'table.items { width: 100%; border-collapse: collapse; margin: 10px 0; }';
  html += 'table.items th { background: #001f3f; color: white; padding: 10px; text-align: left; font-size: 11px; }';
  html += 'table.items td { padding: 8px; border-bottom: 1px solid #ddd; font-size: 11px; }';
  html += 'table.items th:nth-child(1), table.items td:nth-child(1) { text-align: center; width: 5%; }';
  html += 'table.items th:nth-child(5), table.items td:nth-child(5) { text-align: center; }';
  
  // NEW: 2-column layout for Terms & Totals
  html += '.totals-container { display: flex; gap: 30px; margin-top: 30px; }';
  html += '.terms-column { flex: 1; min-width: 300px; }';
  html += '.totals-column { flex: 1; min-width: 350px; }';
  html += '.total-row { display: flex; justify-content: space-between; padding: 8px 0; font-size: 12px; }';
  html += '.total-row span:first-child { text-align: left; min-width: 200px; }';
  html += '.total-row span:last-child { text-align: right; font-weight: 500; }';
  html += '.total-row.section-divider { border-top: 1px solid #ddd; margin-top: 5px; padding-top: 10px; }';
  html += '.total-row.grand-total { font-size: 16px; font-weight: bold; color: #001f3f; border-top: 2px solid #001f3f; margin-top: 5px; padding: 10px 0; }';
  html += '.payment-schedule { background: #f8f9fa; padding: 15px; margin-top: 15px; border-radius: 5px; }';
  html += '.payment-schedule h4 { margin: 0 0 10px 0; font-size: 13px; font-weight: bold; }';
  html += '.footer { clear: both; text-align: center; margin-top: 50px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; }';
  html += '</style></head><body>';

  // Header
  html += '<div class="header">';
  html += '<div class="company-info">';
  html += '<h1>Adams Telecom Systems</h1>';
  html += '<p>Adams Telecom Systems</p>';
  html += '<p>Website: adamstelecom.com</p>';
  html += '<p>Phone: +1 1234567890</p>';
  html += '<p style="margin-top: 10px;">Prepared by: ' + (data.preparedBy || 'N/A') + '</p>';
  html += '</div>';
  html += '<div class="quote-details">';
  html += '<h2>' + title + '</h2>';
  html += '<table class="header-table">';
  html += '<tr><td>' + title + ' #</td><td>' + number + '</td></tr>';
  html += '<tr><td>' + dateLabel + '</td><td>' + dateValue + '</td></tr>';
  html += '<tr><td>' + secondDateLabel + '</td><td>' + secondDateValue + '</td></tr>';
  if (isInvoice && data.quoteNumber) {
    html += '<tr><td>QUOTE #</td><td>' + data.quoteNumber + '</td></tr>';
  }
  html += '</table>';
  html += '</div>';
  html += '</div>';

  // Customer Section
  html += '<div class="section-header">BILL TO</div>';
  html += '<div class="customer-grid">';
  html += '<div class="customer-field"><label>Name</label><div>' + (data.customerName || '') + '</div></div>';
  html += '<div class="customer-field"><label>Company Name</label><div>' + (data.customerCompany || '') + '</div></div>';
  html += '<div class="customer-field"><label>Street Address</label><div>' + (data.customerAddress || '') + '</div></div>';
  html += '<div class="customer-field"><label>City, ST ZIP</label><div>' + (data.customerCity || '') + '</div></div>';
  html += '<div class="customer-field"><label>Phone</label><div>' + (data.customerPhone || '') + '</div></div>';
  html += '<div class="customer-field"><label>Email</label><div>' + (data.customerEmail || '') + '</div></div>';
  html += '</div>';

  // Objective Section (NEW)
  if (data.objective && data.objective.trim() !== '') {
    html += '<div class="objective-box">';
    html += '<h3>PROJECT OBJECTIVE</h3>';
    html += '<p>' + data.objective + '</p>';
    html += '</div>';
  }

  // Items Table (NEW structure)
  html += '<div class="section-header">EQUIPMENT</div>';
  html += '<table class="items">';
  html += '<thead><tr>';
  html += '<th>ITEM</th>';
  html += '<th>MODEL NUMBER</th>';
  html += '<th>ITEM NAME</th>';
  html += '<th>CATEGORY</th>';
  html += '<th>QTY</th>';
  html += '</tr></thead>';
  html += '<tbody>';

  var items = data.items || [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    html += '<tr>';
    html += '<td style="text-align: center;">' + (item.itemNumber || (i + 1)) + '</td>';
    html += '<td>' + (item.modelNumber || '') + '</td>';
    html += '<td>' + (item.itemName || '') + '</td>';
    html += '<td>' + (item.category || '') + '</td>';
    html += '<td style="text-align: center;">' + (item.qty || 0) + '</td>';
    html += '</tr>';
  }

  html += '</tbody></table>';

  // Totals & Terms (2-column layout)
  html += '<div class="totals-container">';
  
  // LEFT COLUMN: Terms & Conditions
  html += '<div class="terms-column">';
  html += '<h4 style="color: #001f3f; margin-bottom: 10px;">TERMS & CONDITIONS</h4>';
  html += '<ul style="font-size: 10px; line-height: 1.6; color: #333; padding-left: 18px; margin: 0;">';
  html += '<li>The recommendations and suggestions made in this proposal are based on information gathered from the provided named areas above and communicated written or verbally from the end user.</li>';
  html += '<li>This information is based on the customer’s request.</li>';
  html += '<li>Adams Telecom Systems makes the recommendations as described in this proposal.</li>';
  html += '<li>Adams Telecom Systems is not liable for any use or information that is not described and/or disclosed in this proposal.</li>';
  html += '</ul>';
  html += '<p></p>' 
  html += '<h4 style="color: #001f3f; margin-bottom: 10px;">Customer Acceptance <span style="font-size:10px;">(Sign below):</span></h4>';
  html += '</div>';
  
  // RIGHT COLUMN: Calculations
  html += '<div class="totals-column">';
  html += '<div class="total-row"><span>Material Cost</span><span>$' + parseFloat(data.materialCost || 0).toFixed(2) + '</span></div>';
  html += '<div class="total-row"><span>Installation & Training</span><span>$' + parseFloat(data.installationCost || 0).toFixed(2) + '</span></div>';
  html += '<div class="total-row section-divider"><span><strong>Sub Total</strong></span><span><strong>$' + parseFloat(data.subTotal || 0).toFixed(2) + '</strong></span></div>';
  // html += '<div class="total-row"><span>Sales Tax</span><span>' + (data.salesTax != null ? '$' + parseFloat(data.salesTax).toFixed(2) : 'Exempt') + '</span></div>';
  html += '<div class="total-row"><span>Sales Tax</span><span>$' + parseFloat(data.salesTax || 0).toFixed(2) + '</span></div>';
  html += '<div class="total-row grand-total"><span>GRAND TOTAL</span><span>$' + parseFloat(data.grandTotal || 0).toFixed(2) + '</span></div>';
  
  // Payment Schedule
  html += '<div class="payment-schedule">';
  //html += '<h4 style="margin: 15px 0 10px 0; color: #001f3f;">PAYMENT SCHEDULE</h4>';
  html += '<h4 style="color: #001f3f;">PAYMENT SCHEDULE</h4>';
  html += '<div class="total-row"><span>Project Down Payment</span><span>$' + parseFloat(data.downPayment || 0).toFixed(2) + '</span></div>';
  html += '<div class="total-row"><span><strong>Final Project Payment</strong></span><span><strong>$' + parseFloat(data.finalPayment || 0).toFixed(2) + '</strong></span></div>';
  html += '</div>';
  
  if (isInvoice) {
    html += '<div class="total-row" style="border-top: 1px solid #ddd; margin-top: 10px;"><span>STATUS</span><span><strong>' + (data.status || 'Unpaid') + '</strong></span></div>';
  }
  html += '</div>'; // Close totals-column
  html += '</div>'; // Close totals-container

  // Footer
  html += '<div class="footer">';
  html += '<p><strong>Thank you for your business!</strong></p>';
  html += '<p style="font-size: 10px; color: #666; margin-top: 10px;">If you have any questions, please contact us at +923224083545 or email us at sales@adamstelecom.com</p>';
  html += '</div>';

  html += '</body></html>';

  return html;
}