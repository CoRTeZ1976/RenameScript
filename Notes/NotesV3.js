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
var docNote;
var currDocID;
var docsCount;
//-----------------------------------------------------------------------------------

//��������� ������� ������ ����������
SApp.StartSelectDocs();
SApp.SelectDocs()
docsCount = SApp.SelectedDocsCount();

//�������� � ����� �����
function createEIObj() {
	oIE = WScript.CreateObject("InternetExplorer.Application", "IE_");
}

function designationForm() {	
	createEIObj();
	oIE.Left = 700;
	oIE.Top = 400;
	oIE.Height = 200;
	oIE.Width = 500;
	oIE.navigate(GetPath() + "form.html");
	oIE.Resizable = false;
	oIE.Visible = 1;
}

//���������� �������� �����
function IE_OnQuit() {
	docNote = oIE.Document.ValidForm.Note.value;
	ready = true;
	oIE.Quit();
}

//��������� ���� � �����
function GetPath() {
	var path = WScript.ScriptFullName;
	path = path.substring(0, path.lastIndexOf("\\") + 1);
	return path;
}

//-----------------------------------------------------------------------------------

SApp.ShowProgressBarForm("������ ����������...", "", "��������", docsCount);

try {
	//�������� �����
	designationForm();
	while (oIE.Busy) {WScript.Sleep(100)};
	ready = false;
	while (!ready) {WScript.Sleep(100)};
	//���������� ��������� ���������
	for (var i = 0; i < docsCount; i++) {
		SApp.SetProgressBarData_and_CheckUserBreak(i, "����������: " + i + " �� " + docsCount, i);
		//�������� ID ���������� ���������
		currDocID = SApp.GetSelectedDocID(i);
		//��������� ��������� ��������
		SApp.OpenDocument(currDocID);
		//����������� ����� ����������
		SApp.SetFieldValue("����������", docNote);
		//������� � �����
		SApp.CheckIn();
	}
	SApp.MessageBox("�������� ������� ���������", "����!", 0);
} catch(e) {
	SApp.MessageBox(e.message, "������", 0);
} finally {
	SApp.CloseProgressBarForm();
	SApp.RefreshCurrentWindow();	
}