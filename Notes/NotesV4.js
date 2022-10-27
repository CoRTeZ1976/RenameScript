Connect2Search()

function Connect2Search() {
	if (typeof(S4App) == "undefined")
{
	SApp = new ActiveXObject("S4.TS4App");
	SApp.Login();
} else
	SApp = S4App;
}

var docNote = "10UMA.MAX.0002"; //<---- ������ ���� ������ ����� ���������� (var docNote = "example")
var currDocID;
var docsCount;
//-----------------------------------------------------------------------------------

SApp.StartSelectDocs();
SApp.SelectDocs()
docsCount = SApp.SelectedDocsCount();


//-----------------------------------------------------------------------------------

SApp.ShowProgressBarForm("������ ����������...", "", "��������", docsCount);

try {
	for (var i = 0; i < docsCount; i++) {
		SApp.SetProgressBarData_and_CheckUserBreak(i, "����������: " + i + " �� " + docsCount, i);
		currDocID = SApp.GetSelectedDocID(i);
		SApp.OpenDocument(currDocID);
		SApp.SetFieldValue("����������", docNote);
		//SApp.CheckIn();
	}
	SApp.MessageBox("�������� ������� ���������", "����!", 0);
} catch(e) {
	SApp.MessageBox(e.message, "������", 0);
} finally {
	SApp.CloseProgressBarForm();
	SApp.RefreshCurrentWindow();	
}