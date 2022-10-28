// add dynamic fields on number of test cases selected
function addFields(evt) {
  // get the value of number of test cases
  if (evt.target.value) {
    // Generate a dynamic number of inputs
    var number = evt.target.value;
    var container = document.getElementById("container");
    // Remove every children it had before
    while (container.hasChildNodes()) {
      container.removeChild(container.lastChild);
    }

    // loop until the number of cases is completed
    // generate number of required input fields
    for (i = 0; i < number; i++) {
      // create a div for each test case
      var mainDiv = document.createElement("div");
      mainDiv.classList.add(`test-cases${i + 1}`);

      // create TC Details Div
      var tcDetails = document.createElement("div");
      tcDetails.classList.add("item");
      var tcDetailsLabel = document.createElement("label");
      tcDetailsLabel.appendChild(document.createTextNode("TC Details: "));
      tcDetailsLabel.style.display = "block";
      tcDetails.appendChild(tcDetailsLabel);
      var tcDetailsTextArea = document.createElement("textArea");
      tcDetailsTextArea.name = "tcDetails" + i;
      tcDetailsTextArea.rows = 5;
      tcDetailsTextArea.cols = 60;
      tcDetailsTextArea.placeholder = "Enter TC Details here ";
      tcDetails.appendChild(tcDetailsTextArea);
      mainDiv.appendChild(tcDetails);

      // create TC name Div
      var tcName = document.createElement("div");
      tcName.classList.add("item");
      var tcNameLabel = document.createElement("label");
      tcNameLabel.appendChild(document.createTextNode("TC Name: "));
      tcNameLabel.style.display = "block";
      tcName.appendChild(tcNameLabel);
      var tcNameInput = document.createElement("input");
      tcNameInput.type = "text";
      tcNameInput.placeholder = "Enter TC Name";
      tcNameInput.name = "tcName" + i;
      tcName.appendChild(tcNameInput);
      mainDiv.appendChild(tcName);

      // create TC Result Div
      var tcResult = document.createElement("div");
      tcResult.classList.add("item");
      var tcResultLabel = document.createElement("label");
      tcResultLabel.appendChild(document.createTextNode("Test Result: "));
      tcResultLabel.style.display = "block";
      tcResult.appendChild(tcResultLabel);
      var tcResultInput = document.createElement("input");
      tcResultInput.type = "text";
      tcResultInput.placeholder = "Enter TC Result";
      tcResultInput.name = "tcResult" + i;
      tcResult.appendChild(tcResultInput);
      mainDiv.appendChild(tcResult);

      // create Expected Behavior Div
      var ExpectedBehavior = document.createElement("div");
      ExpectedBehavior.classList.add("item");
      var ExpectedBehaviorLabel = document.createElement("label");
      ExpectedBehaviorLabel.appendChild(
        document.createTextNode("Expected Behavior: ")
      );
      ExpectedBehaviorLabel.style.display = "block";
      ExpectedBehavior.appendChild(ExpectedBehaviorLabel);
      var ExpectedBehaviorTextArea = document.createElement("textArea");
      ExpectedBehaviorTextArea.name = "ExpectedBehavior" + i;
      ExpectedBehaviorTextArea.rows = 5;
      ExpectedBehaviorTextArea.cols = 60;
      ExpectedBehaviorTextArea.placeholder = "Enter Expected Behavior here";
      ExpectedBehavior.appendChild(ExpectedBehaviorTextArea);
      mainDiv.appendChild(ExpectedBehavior);

      // create Procedure Div
      var Procedure = document.createElement("div");
      Procedure.classList.add("item");
      var ProcedureLabel = document.createElement("label");
      ProcedureLabel.appendChild(document.createTextNode("Procedure: "));
      ProcedureLabel.style.display = "block";
      Procedure.appendChild(ProcedureLabel);
      var ProcedureTextArea = document.createElement("textArea");
      ProcedureTextArea.name = "Procedure" + i;
      ProcedureTextArea.rows = 5;
      ProcedureTextArea.cols = 60;
      ProcedureTextArea.placeholder = "Enter Procedure here";
      Procedure.appendChild(ProcedureTextArea);
      mainDiv.appendChild(Procedure);

      // create image upload div here
      var imageUpload = document.createElement("div");
      imageUpload.classList.add("item");
      var imageUploadLabel = document.createElement("label");
      imageUploadLabel.appendChild(document.createTextNode("Upload Images: "));
      imageUploadLabel.style.display = "block";
      imageUpload.appendChild(imageUploadLabel);
      var imageUploadInput = document.createElement("input");
      imageUploadInput.type = "file";
      imageUploadInput.multiple = true;
      imageUploadInput.placeholder = "Upload Images";
      imageUploadInput.name = "imageUpload" + i;
      imageUpload.appendChild(imageUploadInput);
      mainDiv.appendChild(imageUpload);

      // create Description Div
      var Description = document.createElement("div");
      Description.classList.add("item");
      var DescriptionLabel = document.createElement("label");
      DescriptionLabel.appendChild(document.createTextNode("Description: "));
      DescriptionLabel.style.display = "block";
      Description.appendChild(DescriptionLabel);
      var DescriptionTextArea = document.createElement("textArea");
      DescriptionTextArea.name = "Description" + i;
      DescriptionTextArea.rows = 5;
      DescriptionTextArea.cols = 60;
      DescriptionTextArea.placeholder = "Enter Description here";
      Description.appendChild(DescriptionTextArea);
      mainDiv.appendChild(Description);

      // create Remarks Div
      var Remarks = document.createElement("div");
      Remarks.classList.add("item");
      var RemarksLabel = document.createElement("label");
      RemarksLabel.appendChild(document.createTextNode("Remarks: "));
      RemarksLabel.style.display = "block";
      Remarks.appendChild(RemarksLabel);
      var RemarksTextArea = document.createElement("textArea");
      RemarksTextArea.name = "Remarks" + i;
      RemarksTextArea.rows = 5;
      RemarksTextArea.cols = 60;
      RemarksTextArea.placeholder = "Enter Remarks here";
      Remarks.appendChild(RemarksTextArea);
      mainDiv.appendChild(Remarks);

      // add Test Case no every time
      var testCaseNum = document.createElement("div");
      testCaseNum.classList.add("item");
      var testCaseNumText = document.createTextNode(`Test Case #${i + 1}:`);
      testCaseNum.appendChild(testCaseNumText);

      // Append a line break
      container.appendChild(testCaseNum);
      container.appendChild(mainDiv);
      container.appendChild(document.createElement("br"));
    }
  }
}

// define global variables for document preparation work
let reader = new FileReader();
let uploadImagesBase64 = [];
let msDocData;

// read every image data and convert into base64 format
function readImageData(file) {
  return new Promise((resolve, reject) => {
    try {
      reader.readAsDataURL(file);
      reader.onload = function () {
        uploadImagesBase64.push(reader.result);
        resolve(true);
      };
    } catch (e) {
      reject(e);
    }
  });
}

// on submit form write data MS Word file
async function submitForm(event) {
  // form data object to handle all form data
  var formData = {};
  formData.title = event.target.title.value;
  formData.ticketNumber = event.target.ticketNumber.value;
  formData.ticketOwner = event.target.ticketOwner.value;
  formData.date = event.target.date.value;
  formData.testCases = event.target.testCases.value;

  var testCasesDetails = [];
  // save every test details into test cases details array
  for (let index = 0; index < formData.testCases; index++) {
    testCasesDetails.push({
      tcDetails: event.target["tcDetails" + index].value,
      tcName: event.target["tcName" + index].value,
      tcResult: event.target["tcResult" + index].value,
      ExpectedBehavior: event.target["ExpectedBehavior" + index].value,
      Procedure: event.target["Procedure" + index].value,
      imageUpload: event.target["imageUpload" + index].files,
      Description: event.target["Description" + index].value,
      Remarks: event.target["Remarks" + index].value,
    });
  }

  // loop through each image uploaded by user and convert one by one into base64
  for (let index = 0; index < testCasesDetails.length; index++) {
    if (testCasesDetails[index].imageUpload.length) {
      for (
        let ele = 0;
        ele < testCasesDetails[index].imageUpload.length;
        ele++
      ) {
        await readImageData(testCasesDetails[index].imageUpload[ele]);
      }
      testCasesDetails[index].imageUpload = uploadImagesBase64;
      uploadImagesBase64 = [];
    }
  }

  // add all test cases details into main form data
  formData.testCasesDetails = testCasesDetails;

  // prepare ms word document here using docx package
  msDocData = [
    new docx.Paragraph({
      // title paragraph
      text: `Title: ${formData.title}`,
      heading: "Heading1",
      alignment: "center",
      style: "TitleStyle",
    }),
    new docx.Paragraph({
      // Ticket Number paragraph
      text: "Ticket Number:",
      style: "normalBoldTextHeading",
      children: [
        new docx.TextRun({ text: "\t" + formData.ticketNumber, bold: false }),
      ],
    }),
    new docx.Paragraph({
      // Ticket Owner paragraph
      text: "Ticket Owner:",
      style: "normalBoldTextHeading",
      children: [
        new docx.TextRun({ text: "\t" + formData.ticketOwner, bold: false }),
      ],
    }),
    new docx.Paragraph({
      // Ticket Date paragraph
      text: "Date:",
      style: "normalBoldTextHeading",
      children: [new docx.TextRun({ text: "\t" + formData.date, bold: false })],
    }),
    new docx.Paragraph({
      // Number Of TestCases paragraph
      text: "Number Of TestCases:",
      style: "normalBoldTextHeading",
      children: [
        new docx.TextRun({ text: "\t" + formData.testCases, bold: false }),
      ],
    }),
  ];

  // now loop through all test cases and write every test case in doc
  for (let index = 0; index < formData.testCasesDetails.length; index++) {
    msDocData.push(
      new docx.Paragraph({
        // test case no paragraph
        text: "\n\nTestCase#" + (index + 1) + "\n",
        style: "normalBoldTextHeading",
      })
    );
    msDocData.push(
      new docx.Paragraph({
        // TC Details paragraph
        text: "TC Details:",
        style: "normalBoldTextHeading",
        children: [
          new docx.TextRun({
            text: "\t" + formData.testCasesDetails[index].tcDetails,
            bold: false,
          }),
        ],
      })
    );
    msDocData.push(
      new docx.Paragraph({
        // TC Name paragraph
        text: "TC Name:",
        style: "normalBoldTextHeading",
        children: [
          new docx.TextRun({
            text: "\t" + formData.testCasesDetails[index].tcName,
            bold: false,
          }),
        ],
      })
    );
    msDocData.push(
      new docx.Paragraph({
        // TC Result paragraph
        text: "TC Result:",
        style: "normalBoldTextHeading",
        children: [
          new docx.TextRun({
            text: "\t" + formData.testCasesDetails[index].tcResult,
            bold: false,
          }),
        ],
      })
    );
    msDocData.push(
      new docx.Paragraph({
        // TC Expected Behavior paragraph
        text: "Expected Behavior:",
        style: "normalBoldTextHeading",
        children: [
          new docx.TextRun({
            text: "\t" + formData.testCasesDetails[index].ExpectedBehavior,
            bold: false,
          }),
        ],
      })
    );
    msDocData.push(
      new docx.Paragraph({
        // Procedure paragraph
        text: "Procedure:",
        style: "normalBoldTextHeading",
        children: [
          new docx.TextRun({
            text: "\t" + formData.testCasesDetails[index].Procedure,
            bold: false,
          }),
        ],
      })
    );

    // now loop through every image to be injected into word doc.
    for (
      let item = 0;
      item < formData.testCasesDetails[index].imageUpload.length;
      item++
    ) {
      msDocData.push(
        new docx.Paragraph({
          // every image paragraph
          children: [
            new docx.ImageRun({
              data: formData.testCasesDetails[index].imageUpload[item],
              transformation: {
                width: 200,
                height: 200,
              },
            }),
          ],
        })
      );
    }

    msDocData.push(
      new docx.Paragraph({
        // Description paragraph
        text: "Description:",
        style: "normalBoldTextHeading",
        children: [
          new docx.TextRun({
            text: "\t" + formData.testCasesDetails[index].Description,
            bold: false,
          }),
        ],
      })
    );
    msDocData.push(
      new docx.Paragraph({
        // Remarks paragraph
        text: "Remarks:",
        style: "normalBoldTextHeading",
        children: [
          new docx.TextRun({
            text: "\t" + formData.testCasesDetails[index].Remarks,
            bold: false,
          }),
        ],
      })
    );
  }

  // now write all data into docx api
  const doc = new docx.Document({
    styles: {
      paragraphStyles: [
        {
          id: "TitleStyle",
          name: "Title Style",
          basedOn: "Normal",
          run: {
            bold: true,
            font: "Calibri",
            underline: { type: "single" },
            size: 32,
            shading: {
              color: "#4287f5",
              fill: "#42ddf5",
            },
          },
          paragraph: {
            spacing: { line: 276, before: 150, after: 150 },
          },
        },
        {
          id: "normalBoldTextHeading",
          name: "normal Style",
          basedOn: "Normal",
          run: {
            bold: true,
            font: "Calibri",
          },
        },
      ],
    },
    sections: [
      {
        children: msDocData,
      },
    ],
  });

  // now generate the file
  docx.Packer.toBlob(doc).then((blob) => {
    console.log(blob);
    saveAs(blob, "Report Generator Tool.docx");
    console.log("Document created successfully");
  });

  return true;
}
