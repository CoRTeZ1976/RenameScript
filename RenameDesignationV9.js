Connect2Search()

function Connect2Search() {
	if (typeof(S4App) == "undefined")
{
	SApp = new ActiveXObject("S4.TS4App");
	SApp.Login();
} else
	SApp = S4App;
}

var fileNmae;
var s = "S";
var newFileDes;
var currDocID;
var docsCount;
var docType;
var docName;
var ready;
var isExist;
var isExistDocName = true;

//-----------------------------------------------------------------------------------

SApp.StartSelectDocs();
SApp.SelectDocs();
docsCount = SApp.SelectedDocsCount();
	
//-----------------------------------------------------------------------------------
 
 function getDocData() {
	currDocID = SApp.GetSelectedDocID(i);
	SApp.OpenDocument(currDocID);
	fileName = SApp.GetFieldValue("Имя файла");
	docType = SApp.GetFieldValue("Тип документа");
	isExist = SApp.GetDocID_ByFilename(fileName);
}

function setDesignation(docType) {
	if (docType === "Сканированный чертеж") {
		newFileDes = s.concat(fileName).slice(0, 12);
		SApp.SetFieldValue("Обозначение", newFileDes);
		getWorkDocAndSetNewName(fileName, docType);
	} else if (docType === "Сканированный чертеж сборочный") {
		newFileDes = s.concat(fileName).slice(0, 15);
		SApp.SetFieldValue("Обозначение", newFileDes);
		getWorkDocAndSetNewName(fileName, docType);
	} else if (docType === "Сканированная ДО") {
		newFileDes = s.concat(fileName).slice(0, 15);
		SApp.SetFieldValue("Обозначение", newFileDes);
		getWorkDocAndSetNewName(fileName, docType);
	}
}

function getWorkDocAndSetNewName(fileName, docType) {
	if (docType === "Сканированный чертеж") {
		var workFileName = fileName.slice(0, 11);
	} else if(docType === "Сканированный чертеж сборочный") {
		var workFileName = fileName.slice(0, 14);
	} else if(docType === "Сканированная ДО") {
		var workFileName = fileName.slice(0, 14);
	}
	var currWorkDocId = SApp.GetDocID_ByDesignation(workFileName);
	SApp.OpenDocument(currWorkDocId);
	var workDocName = SApp.GetFieldValue("Наименование");
	if (workDocName === '') {
		isExistDocName = false;
		return;
	} else {
		isExistDocName = true;
	}
	SApp.OpenDocument(currDocID);
	SApp.SetFieldValue("Наименование", workDocName);
}
//----------------------------------------------------------------------------------
	
SApp.ShowProgressBarForm("Заполнение полей...", "", "Прогресс...", docsCount);
try {
	for (var i = 0; i < docsCount; i++) {
		for (var j = i; j < docsCount; j++) {
			SApp.SetProgressBarData_and_CheckUserBreak(i, "Обработано: " + j + " из " + docsCount, j);
			getDocData();
			if (isExist === 0) {
				continue;
			}
			setDesignation(docType);
			i++;
			
			if (isExistDocName === true) {
				SApp.CheckIn();
			}
			
		}
	}
}catch(e) {
	SApp.MessageBox(e.message, "Ошибка", 0);
} finally {
	SApp.CloseProgressBarForm();
	SApp.RefreshCurrentWindow();
	SApp.MessageBox("Операция завершена", "УРААА!", 0);
}
		
		
