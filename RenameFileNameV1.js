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
var currDocID;
var docsCount;
var docName;

//-----------------------------------------------------------------------------------

SApp.StartSelectDocs();
SApp.SelectDocs();
docsCount = SApp.SelectedDocsCount();
	
//-----------------------------------------------------------------------------------
 
 function getDocData() {
	currDocID = SApp.GetSelectedDocID(i);
	SApp.OpenDocument(currDocID);
	docName = SApp.GetFieldValue("�����������") + ".DWG";
	fileNmae = SApp.SetFieldValue("��� �����", docName);
	//var fieldsCount = SApp.GetFieldCount();
	//var fieldName = SApp.GetFieldName(2);
	//SApp.MessageBox(123, SApp.GetFieldValue(fieldName), 0);
	
	//SApp.MessageBox(fieldsCount, fieldName, 0);
	SApp.MessageBox(docName, fileNmae, 0);
}

//----------------------------------------------------------------------------------
	
SApp.ShowProgressBarForm("���������� �����...", "", "��������...", docsCount);
	try {
		for (var i = 0; i < docsCount; i++) {
			for (var j = i; j < docsCount; j++) {
				SApp.SetProgressBarData_and_CheckUserBreak(i, "����������: " + j + " �� " + docsCount, j);
				getDocData();
				i++;
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
		
		
