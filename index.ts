import Excel from "exceljs";
import { appendFileSync, existsSync, unlinkSync } from "fs";

const OutputEmailFile = "output-email.txt";
const OutputPrintFile = "output-print.txt";

type Name = {
  given: string,
  family: string
};

type Member = {
  name: Name,
  apartment: number,
  email: string | null
};

type Electricity = {
  apartment: number,
  consumption: number,
  paidSum: number,
  price: number,
  offset: number
}

function cellValue(row: Excel.Row, cellIndex: number): string {
  return row.getCell(cellIndex).text;
}

function extractMembers(workbook: Excel.Workbook): Member[] {
  let worksheet = workbook.worksheets[0];
  if (!worksheet)  throw "No worksheet found in member workbook";

  // The number was retrieved by checking the excel file
  const StartIndex = 8;

  // But just as a precaution we check that the previous line starts with "Namn"
  const row = worksheet.getRow(StartIndex - 1);
  const value = cellValue(row, 1);
  if (value != "Namn")  throw `StartIndex might be wrong. Line before start was ${value}`;

  // Get all of the rows
  const rows = worksheet.getRows(StartIndex, worksheet.rowCount) ?? [];
  // Filter out all lines that don't have a name
  const nonEmptyRows = rows.filter((row) => cellValue(row, 1));
  // Convert the rows to Member types
  return nonEmptyRows.map((row): Member => {
    // For the name, we split the string by the comma and flip the names
    const name = cellValue(row, 1);
    const nameParts = name.split(",");
    if (nameParts.length != 2)  throw `Name ${name} was malformed to extract name parts`;
    const family = nameParts[0] ?? "";
    const given = nameParts[1] ?? "";

    // The apartment is a long string of the form XX-XXXX-X-YYYY-X of which we are only
    // interested in the YYYY part. So we split by the separating - and get the second to
    // last value in this list
    let fullApartment = cellValue(row, 2).split("-");
    let apartmentStr = fullApartment[fullApartment.length - 2];
    if (!apartmentStr)  throw `No apartment found for ${name}`;

    // The email is pretty straightforward. The email field could be either empty or a
    // - character, both of which we want to interpret as not existing
    let email = cellValue(row, 6);

    return {
      name: { given: given.trim(), family: family.trim() },
      apartment: parseInt(apartmentStr),
      email: (email != "") && (email != "-") ? email : null
    }
  });
}

function extractElectricity(workbook: Excel.Workbook): Electricity[] {
  let worksheet = workbook.worksheets[0];
  if (!worksheet)  throw "No worksheet found in member workbook";

  // The number was retrieved by checking the excel file
  const StartIndex = 3;

  // But just as a precaution we check that the previous line starts with "Lägenhet"
  const row = worksheet.getRow(StartIndex - 1);
  const value = cellValue(row, 1);
  if (value != "Lägenhet")  throw `StartIndex might be wrong. Line before start was ${value}`;

  // Get the electricity price
  let price: number;
  {
    let priceCell = cellValue(worksheet.getRow(1), 6);
    // The price line is of the form: Elkostnad 2023: 2,34 kr/kWh
    // 1. Split by the : to get " 2,34 kr/kWh"
    let s1 = priceCell.split(":")[1];
    if (!s1)  throw "Error in Step 1 of price detection";

    // 2. Trim to get "2,34 kr/kWh"
    let s2 = s1.trim();

    // 3. Split by " " to get "2,34"
    let s3 = s2.split(" ")[0];
    if (!s3)  throw "Error in Step 3 of price detection";

    // 4. Replace "," with "." to get "2.34"
    let s4 = s3.replace(",", ".");

    // 5. Convert to a number
    let s5 = parseFloat(s4);
    if (!s5)  throw "Error in Step 5 of price detection";

    price = s5;
  }

  // Get all of the rows
  const rows = worksheet.getRows(StartIndex, worksheet.rowCount) ?? [];
  // Filter out all lines that don't have an apartment number
  const nonEmptyRows = rows.filter((row) => cellValue(row, 1));
  // Convert the rows to Electricity types
  return nonEmptyRows.map((row): Electricity => {
    const apartment = parseInt(cellValue(row, 1));
    const consumption = parseInt(cellValue(row, 5));
    const paidSum = Math.round(parseFloat(cellValue(row, 7)));
    const offset = Math.round(parseFloat(cellValue(row, 8)));

    return {
      apartment, consumption, price, paidSum, offset
    }
  });
}

function outputMessage(member: Member, electricity: Electricity) {
  console.assert(member.apartment == electricity.apartment);

  if (member.email) {
    let message = `Email: ${member.email}
Subject:  Justering av elkostnad 2022-nov -- 2023-okt


Hej ${member.name.given},
här kommer information om eldebitering för perioden 2022-november tom 2023-oktober för bostadsrätt ${member.apartment}

Förbrukning under perioden: ${electricity.consumption} kWh
Kostnad per kWh under perioden: ${electricity.price.toLocaleString("sv")} kr
El debitering betald under perioden: ${electricity.paidSum} kr
Justering eldebiteringen för er lägenhet att ${electricity.offset > 0 ? "tillägg att betala" : "få åter"}: ${Math.abs(electricity.offset)} kr
Det kommer regleras på er avi för april.

Hälsningar
/ BRF Folkparken Styrelse



`;
    appendFileSync(OutputEmailFile, message);
  }
  else {
    let message = `Hej ${member.name.given},
här kommer information om eldebitering för perioden 2022-november tom 2023-oktober för bostadsrätt ${member.apartment}

Förbrukning under perioden: ${electricity.consumption} kWh
Kostnad per kWh under perioden: ${electricity.price.toLocaleString("sv")} kr
El debitering betald under perioden: ${electricity.paidSum} kr
Justering eldebiteringen för er lägenhet att ${electricity.offset > 0 ? "tillägg att betala" : "få åter"}: ${Math.abs(electricity.offset)} kr
Det kommer regleras på er avi för april.

Hälsningar
/ BRF Folkparken Styrelse



`;
    appendFileSync(OutputPrintFile, message);
  }

}

const main = async () => {
  let memberFile = process.argv[2] ?? "Medlemsregister 2024-01-11 mail.xlsx"
  let electricityFile = process.argv[3] ?? "Brf Folkparken el 2023_till_HSB.xlsx"

  if (existsSync(OutputEmailFile))  unlinkSync(OutputEmailFile);
  if (existsSync(OutputPrintFile))  unlinkSync(OutputPrintFile);

  let membersBook = new Excel.Workbook();
  await membersBook.xlsx.readFile(memberFile);
  let allMembers = extractMembers(membersBook);
  // console.log(members);

  let electricity = new Excel.Workbook();
  await electricity.xlsx.readFile(electricityFile);
  let electricities = extractElectricity(electricity);
  // console.log(electricities);

  let apartmentsWithoutEmail: number[] = [];
  for (let entry of electricities) {
    // First filter out all members by the current apartment number
    let members = allMembers.filter((member) => member.apartment == entry.apartment);

    // Then filter by whether an email is present
    members = members.filter((member) => member.email != null);

    // If not members are left, we have an apartment without email address
    if (members.length == 0) {
      apartmentsWithoutEmail.push(entry.apartment);
      continue;
    }

    // There might be emails that are registered for multiple people, but we only want to
    // send the email once
    let listSent: string[] = [];
    for (let member of members) {
      // The email can't really be empty since we filtered above, but what the heck
      if (member.email == null)  continue;

      // Check for duplicate emails
      if (member.email in listSent)  continue;

      // Add it to the list for the next round
      listSent.push(member.email);
      outputMessage(member, entry);
    }
  }

  for (let apartment of apartmentsWithoutEmail) {
    let members = allMembers.filter((member) => member.apartment == apartment);
    let electricity = electricities.filter((elem) => elem.apartment == apartment);
    console.assert(electricity.length == 1);

    if (members.length == 1) {
      outputMessage(members[0]!, electricity[0]!);
    }
    else {
      // If we have multiple members for a regular message, we just combine their names
      let combinedGiven: string[] = [];
      for (let member of members) {
        combinedGiven.push(member.name.given);
      }
      ;

      let newMember: Member = {
        name: { given: combinedGiven.join(" och "), family: ""},
        apartment: members[0]!.apartment,
        email: null
      };

      outputMessage(newMember, electricity[0]!);
    }
  }
}

main().then();
