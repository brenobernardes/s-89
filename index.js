const reader = require("xlsx");
const { PDFDocument } = require("pdf-lib");
const fileSystem = require("fs");
const path = require("path");

const file = reader.readFile("/path/to/file.xlsx");

let data = [];

const sheets = file.SheetNames;

for (let i = 0; i < sheets.length; i++) {
  const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
  temp.forEach((res) => {
    data.push(res);
    let name = res.Name;
    let assistant = res.Assistant;
    let date = res.Day;
    let assignment = res.Assignment;
    let study = res.Study;
    createPDF(name, assistant, date, assignment, study).catch((err) =>
      console.log(err)
    );
  });
}

async function createPDF(name, assistant, date, assignment, study) {
  const document = await PDFDocument.load(
    fileSystem.readFileSync("/path/to/form/S-89.pdf")
  );

  const form = document.getForm();

  // Declaring the field names on the pdf file

  const studentsName = form.getTextField("900_1_Text");
  const assistantsName = form.getTextField("900_2_Text");
  const assignmentDate = form.getTextField("900_3_Text");
  const bibleReadingCkBox = form.getCheckBox("900_4_CheckBox");
  const initialCallCkBox = form.getCheckBox("900_5_CheckBox");
  const initialCallText = form.getTextField("900_6_Text");
  const returnVisitCkBox = form.getCheckBox("900_7_CheckBox");
  const returnVisitText = form.getTextField("900_8_Text");
  const bibleStudyCkBox = form.getCheckBox("900_9_CheckBox");
  const talkCkBox = form.getCheckBox("900_10_CheckBox");
  const otherCkBox = form.getCheckBox("900_11_CheckBox");
  const otherText = form.getTextField("900_12_Text");
  const mainHall = form.getCheckBox("900_13_CheckBox");
  const bRoom = form.getCheckBox("900_14_CheckBox");
  const cRoom = form.getCheckBox("900_15_CheckBox");

  if (assistant == undefined) {
    setOneStudent();
  } else {
    setTwoStudents();
  }

  function bibleReading() {
    bibleReadingCkBox.check();
  }

  function initialCall() {
    initialCallCkBox.check();
  }

  function initialCallWithStudy() {
    initialCallCkBox.check();
    initialCallText.setText(`${study}`);
  }

  function returnVisit() {
    returnVisitCkBox.check();
  }

  function returnVisitWithStudy() {
    returnVisitCkBox.check();
    returnVisitText.setText(`${study}`);
  }

  function bibleStudy() {
    bibleStudyCkBox.check();
  }

  function talk() {
    talkCkBox.check();
  }

  function setOneStudent() {
    studentsName.setText(`${name}`);
    assignmentDate.setText(`${date}`);

    if (assignment == "BR") {
      bibleReading();
    } else if (assignment == "IC") {
      initialCall();
    } else if (assignment == "RV") {
      returnVisit();
    } else if (assignment == "BS") {
      bibleStudy();
    } else if (assignment == "TK") {
      talk();
    }

    mainHall.check();

    saveFile();
  }

  function setTwoStudents() {
    studentsName.setText(`${name}`);
    assistantsName.setText(`${assistant}`);
    assignmentDate.setText(`${date}`);

    if (assignment == "BR") {
      bibleReading();
    } else if (assignment == "IC" && study != "") {
      initialCallWithStudy();
    } else if (assignment == "IC") {
      initialCall();
    } else if (assignment == "RV" && study != "") {
      returnVisitWithStudy();
    } else if (assignment == "RV") {
      returnVisit();
    } else if (assignment == "BS") {
      bibleStudy();
    } else if (assignment == "TK") {
      talk();
    }

    mainHall.check();

    saveFile();
  }

  async function saveFile() {
    // Formating and saving folder
    const savingDate = new Date();
    const formatedDate = savingDate.toISOString().substr(0, 10);
    const folderName = `${formatedDate}`;

    // Path to save folder (Edited to save on Macbook)
    const desktopPath = path.join(
      require("os").homedir(),
      "Desktop",
      folderName
    );

    // Saving file inside folder
    if ((assistant === "") | (assistant == undefined)) {
      var fileName = `${name}-${assignment}-${study}.pdf`;
      await checkingFileCondition();
    } else {
      var fileName = `${name}-${assistant}.pdf`;
      await checkingFileCondition();
    }

    async function checkingFileCondition() {
      const savingFile = path.join(desktopPath, fileName);

      if (!fileSystem.existsSync(desktopPath)) {
        fileSystem.mkdirSync(desktopPath);
      }
      fileSystem.writeFileSync(savingFile, await document.save());
    }
  }
}
