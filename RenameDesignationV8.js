//������������ � Search
Connect2Search()

function Connect2Search() {
	if (typeof(S4App) == "undefined")
{
	SApp = new ActiveXObject("S4.TS4App");
	SApp.Login();
} else
	SApp = S4App;
}

//��������� ����������
var fileNmae;
var currDocID;
var s = "S";
var newFileDes;
var currDocID;
var docsCount;
var docType;
var docName;
var checkName;
var ready;
var oIE;

//-----------------------------------------------------------------------------------

//��������� ������� ������ ����������
SApp.StartSelectDocs();
SApp.SelectDocs();
docsCount = SApp.SelectedDocsCount();
	
//-----------------------------------------------------------------------------------

//�������� � ����� �����
function createEIObj() {
	oIE = WScript.CreateObject("InternetExplorer.Application", "IE_");

}

function designationForm() {
		
	createEIObj();
	oIE.Left = 700;
	oIE.Top = 400;
	oIE.Height = 350;
	oIE.Width = 580;
	oIE.navigate(GetPath() + "form.html");
	oIE.Resizable = false;
	oIE.Visible = 1;

	//��������� ����� ���� � �������� � �����	
	getDocData();
	var currDraw = oIE.parent.document.getElementsByName('currDraw');
	currDraw[0].innerHTML = fileName;

}

//���������� �������� �����
function IE_OnQuit() {
	//docName = oIE.Document.ValidForm.Name.value;
	var checkBoxName = oIE.parent.document.getElementsByName('checkName');
	if (checkBoxName[0].checked) {
		checkName = true;
	} else {
		checkName = false;
	}
	ready = true;
	oIE.Quit();
}

//��������� ���� � �����
function GetPath() {
	var path = WScript.ScriptFullName;
	path = path.substring(0, path.lastIndexOf("\\") + 1);
	return path;
}
//----------------------------------------------------------------------------------
 
 function getDocData() {
	//�������� ID ���������� ���������
	currDocID = SApp.GetSelectedDocID(i);
	//��������� ��������� �������� (��������)
	SApp.OpenDocument(currDocID);
	//�������� ��� ����� � ��� ���������
	fileName = SApp.GetFieldValue("��� �����");
	docType = SApp.GetFieldValue("��� ���������");
}

function setDesignation(docType) {
	//��������� ��� ���������
	if (docType === "������������� ������") {
		//�������� ����������� � ������������ ����� �� S���-...
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
	SApp.OpenDocument(currDocID);
	SApp.SetFieldValue("������������", workDocName);
}
//----------------------------------------------------------------------------------

//���������� ��������� ���������
	
SApp.ShowProgressBarForm("���������� �����...", "", "��������...", docsCount);
	try {
		for (var i = 0; i < docsCount; i++) {
			if (checkName) {
				for (var j = i; j < docsCount; j++) {
					SApp.SetProgressBarData_and_CheckUserBreak(i, "����������: " + j + " �� " + docsCount, j);
					getDocData();
					setDesignation(docType);
					i++;
					//������� � �����
					SApp.CheckIn();
				}
				break;
			} else {
				SApp.SetProgressBarData_and_CheckUserBreak(i, "����������: " + i + " �� " + docsCount, i);
				getDocData();
				setDesignation(docType);
				SApp.CheckIn();
			}
		}
		SApp.MessageBox("�������� ��������� �������!", "�����!", 0);
	} catch(e) {
		SApp.MessageBox(e.message, "������", 0);
	} finally {
		SApp.CloseProgressBarForm();
		SApp.RefreshCurrentWindow();
	}
		
		
