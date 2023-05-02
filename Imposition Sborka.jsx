#targetengine "session"
var inddColor = { oneSide: 1, twoside: 2, workAndTurn: 4 }; //1 - printing 4+0; 2 - printing 4+4; 4 - printing 4+4 work-and-turn
var CONST = { fileListPDFsName: "ListPDFs.txt", lengthCropMark: 5, offsetCropMark: 1, weightCropMark: 0.25 };
/**
* Main class
*/
var CurrentJob = {
    side1: { listPDFs: [], listPDFs_side2_WAT: [], listBounds: [] }, //WAT - Work-and-Turn
    side2: { listPDFs: [], listBounds: [] },
    imposition: inddColor.twoside,
    isWorkAndTurn: false,
	printHouse: GetPrintHouse(),
    fileListPDFsFullName: decodeURIComponent(app.activeDocument.filePath + "/" + CONST.fileListPDFsName),
    listFromFile: [],
    listMarks: [],
    Initialize: function() {
        with(app.activeDocument.viewPreferences){
            horizontalMeasurementUnits = MeasurementUnits.millimeters;
            verticalMeasurementUnits = MeasurementUnits.millimeters;
        }
        CurrentJob.imposition = CurrentJob.ImpositionFromName();
        CurrentJob.isWorkAndTurn = (CurrentJob.imposition == inddColor.workAndTurn) ? true : false;
        CurrentJob.InddAnalysis();
        CurrentJob.ParseListPDFs();
        CurrentJob.CheckOut();
        CurrentJob.MakeEvent();
    },
    InddAnalysis: function() {
        var pageCenter = Round((app.activeDocument.pages[0].bounds[1] + app.activeDocument.pages[0].bounds[3]) / 2, 2);
        for (var page = 1; page < 3; page++) {
            var rectanglesOnSide = app.activeDocument.pages[page-1].rectangles;
            for (var i = 0; i < rectanglesOnSide.count(); i++) {
                CurrentJob["side" + page].listBounds.push({ x1: Round(rectanglesOnSide[i].geometricBounds[1], 2), y1: Round(rectanglesOnSide[i].geometricBounds[0], 2),
                                                                        x2: Round(rectanglesOnSide[i].geometricBounds[3], 2), y2: Round(rectanglesOnSide[i].geometricBounds[2], 2) });
                if (rectanglesOnSide[i].allGraphics.length > 0)
                    if ((page == 1) && CurrentJob.isWorkAndTurn && (rectanglesOnSide[i].geometricBounds[1] > pageCenter -1))
                        CurrentJob["side" + page].listPDFs_side2_WAT.push(decodeURIComponent(File(rectanglesOnSide[i].allGraphics[0].itemLink.filePath).name));
                    else
                        CurrentJob["side" + page].listPDFs.push(decodeURIComponent(File(rectanglesOnSide[i].allGraphics[0].itemLink.filePath).name));
            }
        }
    },
    ImpositionFromName: function() { //4+0 or 4+4 or WAT
        var color = /(?:0|1|4)\+(0|1|4)(\([cс][oо]\))?/i.exec(app.activeDocument.name);
        if (color.length < 1 ) return false; //name does not contain imposition information
        if (color[2] != undefined) return inddColor.workAndTurn;
        else
            if (color[1] == 0) return inddColor.oneSide;
            else return inddColor.twoside;
    },
    ParseListPDFs: function () { //parsing a file containing a list of files to print
        fileListPDFs = File(CurrentJob.fileListPDFsFullName);
        if (!fileListPDFs.exists) CurrentJob.listFromFile = false;
        else CurrentJob.listFromFile = [];
        fileListPDFs.open('r');
        while (!(fileListPDFs.eof)) {
            var regPDFs = (/(.+)#1#(.+)#2#(\d{1,2})#3#(\d\+\d)#4#(.*)/i.exec(fileListPDFs.readln()));
            if (regPDFs.length > 1)
                    CurrentJob.listFromFile.push({  fileName: regPDFs[1], color: regPDFs[4], postpress: regPDFs[5].replace(/^\s+/, "").replace(/\s+$/, ""), places: regPDFs[3] });
        }
        fileListPDFs.close();
    },
    CountPostPress: function() { //counting orders with postpress
        var count = 0;
        for (var i = 0; i < CurrentJob.listFromFile.length; i++)
            if (CurrentJob.listFromFile[i].postpress != "")
                count += parseInt(CurrentJob.listFromFile[i].places);
        return count;
    },
    MakeEvent: function() { //event at the end of asynchronous actions
        app.addEventListener("afterExport", function(evnt) {
        var task, listener;           
        task = app.idleTasks.add({ name: "exportPDF", sleep: 1000});
        listener = task.addEventListener(IdleEvent.ON_IDLE,  function(ev) {
            listener.remove();
            task.remove();
            NextAsynchronousActions();
            });
        }).name = "exportPDF";
    },
    CheckOut: function() { //checking
        if (CurrentJob.imposition === false) { //incorrect name
            alert("Проверьте название файла", "Некорректно подписан файл...", true);
            exit();
        }
        if ((((CurrentJob.imposition == inddColor.oneSide) || (CurrentJob.imposition == inddColor.workAndTurn)) && (CurrentJob.side2.listPDFs.length != 0)) || //imposition color does not match the color in the file name
            ((CurrentJob.imposition == inddColor.twoside) && (CurrentJob.side2.listPDFs.length == 0))) {
            alert("Проверьте цветность в названии файла", "Неправильная цветность сборки...", true);
            exit();
        }
        if (app.activeDocument.pages.count() != 4) { //incorrect page number
            alert("В файле должно быть 4 страницы", "Файл неправильно подготовлен...", true);
            exit();
        }
        if (CurrentJob.listFromFile !== false) //incorrect number of orders with postpress
            if (CurrentJob.CountPostPress() != CurrentJob.CountPageItemsOnLayer(4, "Schema"))
                if (!confirm("Количество макетов с постобработкой не соответствует количеству, указанному в файле.\n\nПродолжить выполнение программы?", true, "Ошибка в схеме постобработки..."))
                    exit();
        var resultLayoutColorAndSideCount = CurrentJob.LayoutColorAndSideCount();
        if (resultLayoutColorAndSideCount !== true) //page number problem
            if (!confirm(resultLayoutColorAndSideCount + "\n\nПродолжить выполнение программы?", true, "Ошибка цветности или количества сторон..."))
                exit();
        if(CurrentJob.HasIntersectingRectangles()) //misaligned layouts 
            if (!confirm("Некоторые из макетов смещены и налезли на другие.\n\nПродолжить выполнение программы?", true, "Макеты наложились друг на друга..."))
                    exit();
    },
    CountPageItemsOnLayer: function(pageNumber, layerName) { //page items count on the specified layer
        count = 0;
        for (var i = 0; i < app.activeDocument.pages[pageNumber-1].pageItems.count(); i++)
            if (app.activeDocument.pages[pageNumber-1].pageItems[i].itemLayer.name == layerName)
                count ++;
        return count;
    },
    LayoutColorAndSideCount: function() { //
        var list = {};
        var result = "";
        listPDFsSide1 = CurrentJob.side1.listPDFs;
        listPDFsSide2 = (CurrentJob.isWorkAndTurn) ? CurrentJob.side1.listPDFs_side2_WAT : CurrentJob.side2.listPDFs;
        for (var i = 0; i < CurrentJob.listFromFile.length; i++)
            list[CurrentJob.listFromFile[i].fileName] = { side1Task: CurrentJob.listFromFile[i].places, side2Task: (CurrentJob.listFromFile[i].color == "4+0") ? 0 : CurrentJob.listFromFile[i].places,
                                       side1Fact: 0, side2Fact: 0 };
        for (var fileName in listPDFsSide1)
            if (list[listPDFsSide1[fileName]] != undefined) list[listPDFsSide1[fileName]].side1Fact += 1;
        for (var fileName in listPDFsSide2)
            if (list[listPDFsSide2[fileName]] != undefined) list[listPDFsSide2[fileName]].side2Fact += 1;
        for (var fileName in list)
            if ((list[fileName].side1Task !=list[fileName].side1Fact) || (list[fileName].side2Task !=list[fileName].side2Fact)) {
                if (result == "") result = "Неправильная цветность и/или количество макетов:";
                result += ("\n\n" + fileName + "\n" + list[fileName].side1Task + " - " + list[fileName].side2Task + "\n" + list[fileName].side1Fact + " - " + list[fileName].side2Fact);
            }
        return (result == "") ? true : result;
    },
    HasIntersectingRectangles: function() { //misaligned layouts 
        for (var i = 0; i<CurrentJob["side1"].listBounds.length; i++)
            for (var j = i+1; j<CurrentJob["side1"].listBounds.length; j++) {
                if(RectanglesIntersect(CurrentJob["side1"].listBounds[i], CurrentJob["side1"].listBounds[j]))
                    return true;
            }
        return false;
    }
}

//list of operations to be carried out for each type of imposition
var Actions = [{ label: "Удаление слоя с заготовками схем постпресса", Task: function() { RemoveLayer("inactiveSchema") }, dialogLabel: null, //1
				  imposition: inddColor.oneSide | inddColor.twoside | inddColor.workAndTurn },
			   { label: "Форматирование подписей для превью", Task: FormatServiceInformation, dialogLabel: null, //2
				  imposition: inddColor.oneSide | inddColor.twoside | inddColor.workAndTurn },
			   { label: "Рассчёт координат меток реза на лице", Task: function() { CalculateMarks(1) }, dialogLabel: null, //3
				  imposition: inddColor.oneSide | inddColor.twoside | inddColor.workAndTurn },
			   { label: "Расстановка меток реза на лице", Task: function() { MakeMarks(1) }, dialogLabel: null, //4
				  imposition: inddColor.oneSide | inddColor.twoside | inddColor.workAndTurn },
			   { label: "Рассчёт координат меток реза на обороте", Task: function() { CalculateMarks(2) }, dialogLabel: null, //5
				  imposition: inddColor.twoside },
			   { label: "Расстановка меток реза на обороте", Task: function() { MakeMarks(2) }, dialogLabel: null, //6
				  imposition: inddColor.twoside },
			   { label: "Удаление оборота", Task: function() { RemovePage(2) }, dialogLabel: null, //7
				  imposition: inddColor.oneSide | inddColor.workAndTurn },
			   { label: "Сохранение", Task: function() { SaveDocument() }, dialogLabel: null, //8
				  imposition: inddColor.oneSide | inddColor.twoside | inddColor.workAndTurn },
			   { label: "Удаление файла со списком PDF'ок", Task: function() { RemoveFile(CurrentJob.fileListPDFsFullName) }, dialogLabel: null, //9
				  imposition: inddColor.oneSide | inddColor.twoside | inddColor.workAndTurn }];
var AsynchronousActions = { currentAction: 0, listActions: GetListOfAsyncActions()};
                      

var time = Date.now(); //current date

function GetListOfAsyncActions() //asynchronous tasks, saving to PDF
{
	var listActions = [];
	if (CurrentJob.printHouse == "Roteks") { //dependence on the production shop
		listActions.push(function() { AsynchronousToPDF(CurrentJob.printHouse, "", (CurrentJob.imposition == inddColor.oneSide || CurrentJob.imposition == inddColor.workAndTurn) ? "1" : "1-2"); });
		listActions.push(function() { AsynchronousToPDF("Preview", "Preview_", PageRange.ALL_PAGES); });
	} else {
		listActions.push(function() { AsynchronousToPDF(CurrentJob.printHouse, "", PageRange.ALL_PAGES); });
	}
	return listActions;
}


CurrentJob.Initialize(); //start
var StatusDialog = new Window ('palette', 'Sborka.ua'); //dialog
StatusDialog.spacing = 5;
StatusDialog.preferredSize = [340, -1];
for (var index in Actions) //show list of actions
     if (CurrentJob.imposition & Actions[index].imposition) {
        Actions[index].dialogLabel = StatusDialog.add('staticText', [20, 0, 300, 17], Actions[index].label);
        Actions[index].dialogLabel.graphics.foregroundColor = Actions[index].dialogLabel.graphics.newPen (Actions[index].dialogLabel.graphics.PenType.SOLID_COLOR, [0.6, 0.6, 0.6], 1);
}
StatusDialog.ProgressBar = StatusDialog.add('progressbar', [0, 0, 300, 15], 0, 100);
StatusDialog.show(); //show dialog
for (var index in Actions) //UI upgrading, show current task
    if (CurrentJob.imposition & Actions[index].imposition) {
        Actions[index].dialogLabel.graphics.foregroundColor = Actions[index].dialogLabel.graphics.newPen (Actions[index].dialogLabel.graphics.PenType.SOLID_COLOR, [0, 0, 1], 1);
        Actions[index].Task();
        $.sleep(250);
        Actions[index].dialogLabel.graphics.foregroundColor = Actions[index].dialogLabel.graphics.newPen (Actions[index].dialogLabel.graphics.PenType.SOLID_COLOR, [0.1, 0.1, 0.1], 1);
    }
StatusDialog.close();
NextAsynchronousActions(); //publishing to PDF

function NextAsynchronousActions() { //dialog of asynchronous tasks
    if ((AsynchronousActions.currentAction == 0) && (AsynchronousActions.currentAction < AsynchronousActions.listActions.length))
        app.panels.item("Background Tasks").visible = true;
    if (AsynchronousActions.currentAction < AsynchronousActions.listActions.length) {
        AsynchronousActions.currentAction++;
        AsynchronousActions.listActions[AsynchronousActions.currentAction - 1]();
    }
    else {
        app.panels.item("Background Tasks").visible = false;
    myListener = app.eventListeners.itemByName("exportPDF");
    if (myListener.isValid)
        myListener.remove();
    app.activeDocument.close(SaveOptions.NO); //closing the document when all tasks are completed
    }
}

function AsynchronousToPDF(presetName, prefixFileName, pageRange){ //publishing to PDF
    var PDFExportPreset = app.pdfExportPresets.item(presetName);
    if (!PDFExportPreset.isValid) {
        alert("Не удалось загрузить ПДФ-настройку \"" + presetName + "\"");
        NextAsynchronousActions();
        return;
        }
    app.pdfExportPreferences.pageRange = pageRange;
    var PDFName = decodeURIComponent(app.activeDocument.filePath + "/" + prefixFileName + app.activeDocument.name);
    PDFName = PDFName.substr(0, PDFName.lastIndexOf('.'))  + ".pdf";
    app.activeDocument.asynchronousExportFile(ExportFormat.pdfType, File(PDFName), false, PDFExportPreset);
}

function SaveDocument() { //saving the document
    ZoomFirstPage();
	var fileName = decodeURIComponent(app.activeDocument.filePath + "/" + app.activeDocument.name);
    app.activeDocument.save(new File(fileName));
}

function RemoveFile(path) { //removing the file with a list of layouts 
    try {
        File(path).remove();
    }
    catch (myError) { };
}

function RemovePage(pageNumber) { //removing a page
    app.activeDocument.pages[pageNumber - 1].remove();
}

function RemoveLayer(layerName) { //removing a layer
    if (app.activeDocument.layers.item(layerName).isValid)
        app.activeDocument.layers.item(layerName).remove();
}

function FormatServiceInformation() { //formatting technical information 
    ZoomFirstPage();
	try {
		for (var i = 0; i < app.activeDocument.masterSpreads[0].pageItems.count(); i++) {
			if ((app.activeDocument.masterSpreads[0].pageItems[i].geometricBounds[0] < 0) && (app.activeDocument.paragraphStyles.item("Top").isValid))
				app.activeDocument.masterSpreads[0].textFrames[i].parentStory.paragraphs[0].appliedParagraphStyle = app.activeDocument.paragraphStyles.item("Top");
			if ((app.activeDocument.masterSpreads[0].pageItems[i].geometricBounds[0] > 0) && (app.activeDocument.paragraphStyles.item("Buttom").isValid))
				app.activeDocument.masterSpreads[0].textFrames[i].parentStory.paragraphs[0].appliedParagraphStyle = app.activeDocument.paragraphStyles.item("Buttom");
		}
	}
	catch(err) {}
}

function ZoomFirstPage() { //page activation 
    app.layoutWindows[0].activePage = app.activeDocument.pages[0];
    app.layoutWindows[0].zoom(ZoomOptions.FIT_SPREAD);
    app.activeDocument.activeLayer = app.activeDocument.layers.lastItem();
}

function CalculateMarks(pageNumber) { //calculation of cut marks
    ProgressBar = StatusDialog.ProgressBar;
    ProgressBar.value = 0;
    var document = app.activeDocument;
    var page = document.pages[pageNumber - 1];
    app.layoutWindows[0].activePage = page;
    app.layoutWindows[0].zoom(ZoomOptions.FIT_SPREAD);
    var rectangles = CurrentJob["side" + pageNumber].listBounds;
    var listMarks = [];
    ProgressBar.maxvalue = rectangles.length;
    for (var index in rectangles) {
        ProgressBar.value = index;
        for (var X = 1; X < 3; X++)
            for (var Y = 1; Y < 3; Y++) {
                var vectorOffsetX  = ((X%2 == 0) ? -1 : 1);
                var vectorOffsetY  = ((Y%2 == 0) ? -1 : 1);
                var mark = { };
                mark.x1 = mark.x2 = rectangles[index]["x" + X] + vectorOffsetX * CONST.offsetCropMark;
                mark.y1 = rectangles[index]["y" + Y];
                mark.y2 = rectangles[index]["y" + Y] - vectorOffsetY * CONST.lengthCropMark;
                if (CheckoutMark(mark, pageNumber))
                    listMarks.push([mark.y1, mark.x1, mark.y2, mark.x2])
                else // short mark if there is not enough space for long ones
                    if (CONST.lengthCropMark > 2) {
                        mark.y2 = rectangles[index]["y" + Y] - vectorOffsetY * (CONST.lengthCropMark - 1);
                        if (CheckoutMark(mark, pageNumber)) {
                            mark.y2 = rectangles[index]["y" + Y] - vectorOffsetY * (CONST.lengthCropMark - 2.5);
                            listMarks.push([mark.y1, mark.x1, mark.y2, mark.x2]);
                        }
                    }
                var mark = { };
                mark.x1 = rectangles[index]["x" + X];
                mark.x2 = rectangles[index]["x" + X] - vectorOffsetX * CONST.lengthCropMark;
                mark.y1 = mark.y2 = rectangles[index]["y" + Y] + vectorOffsetY * CONST.offsetCropMark;
                if (CheckoutMark(mark, pageNumber))
                    listMarks.push([mark.y1, mark.x1, mark.y2, mark.x2])
                else // short mark if there is not enough space for long ones
                    if (CONST.lengthCropMark > 2) {
                        mark.x2 = rectangles[index]["x" + X] - vectorOffsetX * (CONST.lengthCropMark - 1);
                        if (CheckoutMark(mark, pageNumber)) {
                            mark.x2 = rectangles[index]["x" + X] - vectorOffsetX * (CONST.lengthCropMark - 2.5);
                            listMarks.push([mark.y1, mark.x1, mark.y2, mark.x2]);
                        }
                    }
            }
    }
    CurrentJob.listMarks = listMarks;
    ProgressBar.value = 0;
    return listMarks;
}

function CheckoutMark(mark, pageNumber) { //checking the possibility of adding a cut mark
    var rectangles = CurrentJob["side" + pageNumber].listBounds;
    for (var index = 0; index < rectangles.length; index++)
        if ((mark.x2 >= rectangles[index].x1) && (mark.x2 <= rectangles[index].x2) && (mark.y2 >= rectangles[index].y1) && (mark.y2 <= rectangles[index].y2))
            return false;
    return true;
}

function MakeMarks(pageNumber) { //marks rendering
    ProgressBar = StatusDialog.ProgressBar;
    ProgressBar.value = 0;
    var listMarks = CurrentJob.listMarks;
    var document = app.activeDocument;
    var page = document.pages[pageNumber-1];
    var layerForMarks = (document.layers.item("Marks").isValid) ? document.layers.item("Marks") : document.activeLayer;
    ProgressBar.maxvalue = listMarks.length;
    for (var mark in listMarks) {
        var lineMark = page.graphicLines.add(layerForMarks);
        lineMark.geometricBounds = listMarks[mark];
        lineMark.strokeWeight = CONST.weightCropMark + "pt";
        lineMark.strokeAlignment = 1936998723;
        lineMark.rightLineEnd = 1852796517;
        lineMark.leftLineEnd = 1852796517;
        lineMark.endCap = 1650680176;
        lineMark.endJoin = 1835691886;
        lineMark.strokeColor = document.swatches.item("Registration");
        lineMark.overprintStroke = true;
        lineMark.nonprinting = false;
        lineMark.strokeType = document.strokeStyles.item("Solid");
        ProgressBar.value = mark;
    }
    ProgressBar.value = 0;
}

function Round(value, digits)
{
    return Math.round(value * Math.pow(10, digits)) / Math.pow(10, digits);
}

function RectanglesIntersect(rect1, rect2)
{
    return Math.max(rect1.x1, rect2.x1) < Math.min(rect1.x2, rect2.x2) && Math.max(rect1.y1, rect2.y1) < Math.min(rect1.y2, rect2.y2);
}

function GetPrintHouse() //production shop from name
{
	var printHouse = /Adworld|Roteks/i.exec(app.activeDocument.name);
	return (printHouse == null ) ? "Wolf" : printHouse[0];
}