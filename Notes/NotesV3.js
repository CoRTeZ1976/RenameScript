//Подключаемся к Search
Connect2Search()

function Connect2Search() {
	if (typeof(S4App) == "undefined")
{
	SApp = new ActiveXObject("S4.TS4App");
	SApp.Login();
} else
	SApp = S4App;
}

//Объявляем переменные
var docNote;
var currDocID;
var docsCount;
//-----------------------------------------------------------------------------------

//Запускаем процесс выбора документов
SApp.StartSelectDocs();
SApp.SelectDocs()
docsCount = SApp.SelectedDocsCount();

//Создание и вывод формы
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

//Обработчик закрытия формы
function IE_OnQuit() {
	docNote = oIE.Document.ValidForm.Note.value;
	ready = true;
	oIE.Quit();
}

//Получение пути к форме
function GetPath() {
	var path = WScript.ScriptFullName;
	path = path.substring(0, path.lastIndexOf("\\") + 1);
	return path;
}

//-----------------------------------------------------------------------------------

SApp.ShowProgressBarForm("Запись примечаний...", "", "Прогресс", docsCount);

try {
	//Вызываем форму
	designationForm();
	while (oIE.Busy) {WScript.Sleep(100)};
	ready = false;
	while (!ready) {WScript.Sleep(100)};
	//Перебираем выбранные документы
	for (var i = 0; i < docsCount; i++) {
		SApp.SetProgressBarData_and_CheckUserBreak(i, "Обработано: " + i + " из " + docsCount, i);
		//Получаем ID выбранного документа
		currDocID = SApp.GetSelectedDocID(i);
		//Открываем выбранный документ
		SApp.OpenDocument(currDocID);
		//Присваеваем новое примечание
		SApp.SetFieldValue("Примечание", docNote);
		//Возврат в архив
		SApp.CheckIn();
	}
	SApp.MessageBox("Операция успешно завершена", "Ураа!", 0);
} catch(e) {
	SApp.MessageBox(e.message, "Ошибка", 0);
} finally {
	SApp.CloseProgressBarForm();
	SApp.RefreshCurrentWindow();	
}