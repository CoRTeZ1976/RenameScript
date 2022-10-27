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
	fileName = SApp.GetFieldValue("��� �����");
	docType = SApp.GetFieldValue("��� ���������");
	isExist = SApp.GetDocID_ByFilename(fileName);
}

function setDesignation(docType) {
	if (docType === "������������� ������") {
		newFileDes = s.concat(fileName).slice(0, 12);
		SApp.SetFieldValue("�����������", newFileDes);
		getWorkDocAndSetNewName(fileName, docType);
	} else if (docType === "������������� ������ ���������") {
		newFileDes = s.concat(fileName).slice(0, 15);
		SApp.SetFieldValue("�����������", newFileDes);
		getWorkDocAndSetNewName(fileName, docType);
	} else if (docType === "������������� ��") {
		newFileDes = s.concat(fileName).slice(0, 15);
		SApp.SetFieldValue("�����������", newFileDes);
		getWorkDocAndSetNewName(fileName, docType);
	}
}

function getWorkDocAndSetNewName(fileName, docType) {
	if (docType === "������������� ������") {
		var workFileName = fileName.slice(0, 11);
	} else if(docType === "������������� ������ ���������") {
		var workFileName = fileName.slice(0, 14);
	} else if(docType === "������������� ��") {
		var workFileName = fileName.slice(0, 14);
	}
	var currWorkDocId = SApp.GetDocID_ByDesignation(workFileName);
	SApp.OpenDocument(currWorkDocId);
	var workDocName = SApp.GetFieldValue("������������");
	if (workDocName === '') {
		isExistDocName = false;
		return;
	} else {
		isExistDocName = true;
	}
	SApp.OpenDocument(currDocID);
	SApp.SetFieldValue("������������", workDocName);
}
//----------------------------------------------------------------------------------
	
SApp.ShowProgressBarForm("���������� �����...", "", "��������...", docsCount);
try {
	for (var i = 0; i < docsCount; i++) {
		for (var j = i; j < docsCount; j++) {
			SApp.SetProgressBarData_and_CheckUserBreak(i, "����������: " + j + " �� " + docsCount, j);
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
	SApp.MessageBox(e.message, "������", 0);
} finally {
	SApp.CloseProgressBarForm();
	SApp.RefreshCurrentWindow();
	SApp.MessageBox("�������� ���������", "�����!", 0);
}
		
		
