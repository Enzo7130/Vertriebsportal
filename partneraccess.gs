function isActiveSentiPartner_(email) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Senti Partner');

  if (!sheet) return false;

  const partners = sheet.getDataRange().getValues().slice(1);

  return partners.some(row =>
    normalizeEmail_(row[0]) === normalizeEmail_(email)
  );
}
