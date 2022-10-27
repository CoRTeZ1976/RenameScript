Connect2Search()

function Connect2Search() {
	if (typeof(S4App) == "undefined")
{
	SApp = new ActiveXObject("S4.TS4App");
	SApp.Login();
} else
	SApp = S4App;
}

var docNote = "10UMA.MAX.0002"; //<---- Ввести сюда нужный текст примечания (var docNote = "example")
var currDocID;
var docsCount;
//-----------------------------------------------------------------------------------

SApp.StartSelectDocs();
SApp.SelectDocs()
docsCount = SApp.SelectedDocsCount();


//-----------------------------------------------------------------------------------

SApp.ShowProgressBarForm("Запись примечаний...", "", "Прогресс", docsCount);

try {
	for (var i = 0; i < docsCount; i++) {
		SApp.SetProgressBarData_and_CheckUserBreak(i, "Обработано: " + i + " из " + docsCount, i);
		currDocID = SApp.GetSelectedDocID(i);
		SApp.OpenDocument(currDocID);
		SApp.SetFieldValue("Примечание", docNote);
		//SApp.CheckIn();
	}
	SApp.MessageBox("Операция успешно завершена", "Ураа!", 0);
} catch(e) {
	SApp.MessageBox(e.message, "Ошибка", 0);
} finally {
	SApp.CloseProgressBarForm();
	SApp.RefreshCurrentWindow();	
}