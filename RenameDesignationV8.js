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

//Запускаем процесс выбора документов
SApp.StartSelectDocs();
SApp.SelectDocs();
docsCount = SApp.SelectedDocsCount();
	
//-----------------------------------------------------------------------------------

//Создание и вывод формы
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

	//Получение имени фала и передача в форму	
	getDocData();
	var currDraw = oIE.parent.document.getElementsByName('currDraw');
	currDraw[0].innerHTML = fileName;

}

//Обработчик закрытия формы
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

//получение пути к форме
function GetPath() {
	var path = WScript.ScriptFullName;
	path = path.substring(0, path.lastIndexOf("\\") + 1);
	return path;
}
//----------------------------------------------------------------------------------
 
 function getDocData() {
	//Получаем ID выбранного документа
	currDocID = SApp.GetSelectedDocID(i);
	//Открываем выбранный документ (карточка)
	SApp.OpenDocument(currDocID);
	//Получаем имя файла и тип документа
	fileName = SApp.GetFieldValue("Имя файла");
	docType = SApp.GetFieldValue("Тип документа");
}

function setDesignation(docType) {
	//Проверяем тип документа
	if (docType === "Сканированный чертеж") {
		//Изменяем обозначение и наименование файла на SТММ-...
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
	SApp.OpenDocument(currDocID);
	SApp.SetFieldValue("Наименование", workDocName);
}
//----------------------------------------------------------------------------------

//Перебираем выбранные документы
	
SApp.ShowProgressBarForm("Заполнение полей...", "", "Прогресс...", docsCount);
	try {
		for (var i = 0; i < docsCount; i++) {
			if (checkName) {
				for (var j = i; j < docsCount; j++) {
					SApp.SetProgressBarData_and_CheckUserBreak(i, "Обработано: " + j + " из " + docsCount, j);
					getDocData();
					setDesignation(docType);
					i++;
					//Возврат в архив
					SApp.CheckIn();
				}
				break;
			} else {
				SApp.SetProgressBarData_and_CheckUserBreak(i, "Обработано: " + i + " из " + docsCount, i);
				getDocData();
				setDesignation(docType);
				SApp.CheckIn();
			}
		}
		SApp.MessageBox("Операция выполнена успешно!", "УРААА!", 0);
	} catch(e) {
		SApp.MessageBox(e.message, "Ошибка", 0);
	} finally {
		SApp.CloseProgressBarForm();
		SApp.RefreshCurrentWindow();
	}
		
		
