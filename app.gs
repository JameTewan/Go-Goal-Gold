// Code.gs - Apps Script Backend

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips");
    if (!sheet)
      return responseJSON({
        status: "error",
        message: "Sheet 'Trips' not found",
      });

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return responseJSON({ status: "success", data: [] });

    data.shift(); // ลบ Headers

    const trips = data.map((row) => ({
      id: row[0],
      title: row[1],
      destination: row[2],
      category: String(row[3]).toLowerCase(),
      days: Number(row[4]) || 1,
      nights: Number(row[5]) || 0,
      pricePerPerson: Number(row[6]) || 0,
      imageUrl: row[7] || "",
      description: row[8] || "",
      bookingUrl: row[9] || "",
      notes: row[10] || "", // <-- เพิ่มคอลัมน์ K: หมายเหตุ (optional)
    }));

    return responseJSON({ status: "success", data: trips });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requests");
    if (!sheet)
      return responseJSON({
        status: "error",
        message: "Sheet 'Requests' not found",
      });

    const body = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      body.name || "",
      body.email || "",
      body.destination || "",
      body.startDate || "",
      body.travelers || 1,
      body.budgetPerPerson || 0,
      body.notes || "",
    ]);

    return responseJSON({
      status: "success",
      message: "Request recorded successfully",
    });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

function responseJSON(json) {
  return ContentService.createTextOutput(JSON.stringify(json)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
