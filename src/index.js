const { google } = require("googleapis");

//const keys = require('../keys.json');
const keys = require("../credentials.json");

const client = new google.auth.JWT(keys.client_email, null, keys.private_key, [
  "https://www.googleapis.com/auth/spreadsheets",
]);

client.authorize(function (error) {
  if (error) {
    console.log(error);
    return;
  } else {
    console.log("Connected!!");
    gsrun(client);
  }
});

let updateOptions;

const totalAbsences = 60 * 0.25;
const situation = {
  failedLack: "Reprovado por Falta",
  approved: "Aprovado",
  finalExam: "Exame Final",
  disapproved: "Reprovado por Nota",
};

// google sheets run
async function gsrun(cl) {
  // Instance of Google Sheets API
  // keys
  //const spreadsheetId = "1y4tp_sAQqVd6q0Pt5r-z1IatE8YYEzpXeBCSShY0TyY";
  // credentials
  const spreadsheetId = "1ynAyFUEpaWFgHXrESTrUBjwo7PcL9VlUjit_3ssah7s";
  const googleSheets = google.sheets({ version: "v4", auth: cl });

  // Read rows from spreadsheet
  const opt = {
    spreadsheetId,
    range: "tuntsApi!C4:H",
  };
  let data = await googleSheets.spreadsheets.values.get(opt);
  let dataArray = data.data.values;

  console.log("Data: ", data.data);
  console.log("Data length", dataArray.length);

  dataArray.forEach((item, index) => {
    let average = 0;
    let absentStudent;
    let verifyLine;
    let lineCounter = index;
    console.log("Linha: ", lineCounter);

    for (let i = 0; i < item.length; i++) {
      /* 
        Atualmente tenho array da coluna C até H, e cada linha possui 5 posições no array.
        Logo, o array que estiver preenchido na posição 4, não precisa ser preenchido novamente.
        Caso não tenha nada, o resultado vai ser indefinido, pronto para ser preenchido no caso.
      */
      verifyLine = item[4];
      console.log("verifyLine", verifyLine);
      if (verifyLine === undefined) {
        if (i === 0) {
          // faltas do estudante
          absentStudent = item[i];
          console.log(`faltas: ${absentStudent}`);
        }
        if (i !== 0) {
          // media do estudante
          average = average + Number(item[i]);
        }
      }
    }

    // preencher somente se a linha G4 estiver vazia. Subentendo que a H também esteja
    if (verifyLine === undefined) {
      console.log(`media: ${average}`);
      // escrever se ta aprovado ou nao e etc..
      studentSituation(average, absentStudent, lineCounter);
    }
  });

  function studentSituation(mean, lack, lineCounter) {
    // arredondar
    let average = Math.round(mean / 3);

    switch (true) {
      case lack > totalAbsences:
        writeOnSpreadsheet(situation.failedLack, 0, lineCounter);
        break;
      case average >= 70:
        writeOnSpreadsheet(situation.approved, 0, lineCounter);
        break;
      case average >= 50 && average < 70:
        let naf = average - 10;
        writeOnSpreadsheet(situation.finalExam, naf, lineCounter);
        break;
      case average < 50:
        writeOnSpreadsheet(situation.disapproved, 0, lineCounter);
        break;
      default:
        console.log("Operation Invalid!");
    }
  }

  // Write row(s) to spreadsheet (update)
  async function writeOnSpreadsheet(
    resultSituation,
    finalApprovalResult,
    lineCounter
  ) {
    console.log(`result Situation: ${resultSituation}`);
    console.log(`final Approval Result: ${finalApprovalResult}`);

    updateOptions = {
      spreadsheetId,
      // registrar a partir da coluna G na linha 4.
      range: `tuntsApi!G${4 + lineCounter}:H`,
      valueInputOption: "USER_ENTERED",
      resource: {
        values: [[resultSituation, finalApprovalResult]],
      },
    };

    let res = await googleSheets.spreadsheets.values.update(updateOptions);
    console.log({ res });
  }
}
