var objApp = window.external;
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
var objDatabase = objApp.Database;


function setParamValue(doc, key, value) {
    if (!doc) {
        return;
    }

    doc.SetParamValue(key, value);
}


function stringTrim(str) {
    if (!str) {
        return "";
    }

    return str.replace(/(^\s*)|(\s*$)/g, "");
}


function stripHTML(oldString) {
    var newString = "";
    var inTag = false;
    for (var i = 0; i < oldString.length; i++) {
        if (oldString.charAt(i) == '<')
            inTag = true;
        if (oldString.charAt(i) == '>') {
            if (oldString.charAt(i + 1) == "<") {
                //dont do anything
            } else {
                inTag = false;
                i++;
            }

        }
        if (!inTag)
            newString += oldString.charAt(i);
    }

    newString = stringTrim(newString);
    newString = newString.replace(/\&nbsp\;/g, " ");
    return newString;
}


function getTitle(parthtml) {
    return 'title';
}

function getDetailText(parthtml) {
    var detailsText = getTexFromTag(parthtml, "div", "class", "graytext timesep");
    if (null == detailsText)
        return null;
    return detailsText;
}

function stringToDate(text) {
    var newString = "";
    var inTag = false;
    for (var i = 0; i < text.length; i++) {
        var ch = text.charAt(i);
        if (ch == '年' || ch == '月' || ch == '日' || ch == ':') {
            ch = ':';
        } else if (ch >= '0' && ch <= '9') {} else {
            ch = null;
        }
        if (ch == null)
            continue;
        newString += ch;
    }
    //
    var arr = newString.split(":");
    if (arr.length != 5)
        return null;
    //
    try {
        var date = new Date(arr[0], arr[1] - 1, arr[2], arr[3], arr[4], 0);
        //
        return date;
    } catch (err) {
        return null;
    }
}

function getDetail(parthtml) {
    var text = getDetailText(parthtml);
    //
    var arr = text.split("|");
    //
    var detail = {};
    //
    for (var i = 0; i < arr.length; i++) {
        var att = arr[i];
        att = stringTrim(att);
        //
        if (att.indexOf("创建时间：") == 0) {
            var date = stringToDate(att);
            detail["date"] = date;
        } else if (att.indexOf("分类：") == 0) {
            detail["category"] = att.substring(3);
        } else if (att.indexOf("天气：") == 0) {
            detail["weather"] = att;
        }
    }
    //
    return detail;
}

function formatInt(val) {
    if (val < 10)
        return "0" + val;
    else
        return "" + val;
}


function DateToStr(dt) {
    return "" + dt.getFullYear() + "-" + formatInt(dt.getMonth() + 1) + "-" + formatInt(dt.getDate());
}

function DateToChineseStr(dt) {
    return "" + dt.getFullYear() + "年" + formatInt(dt.getMonth() + 1) + "月" + formatInt(dt.getDate()) + "日";
}

function getLocalString(filename, stringName) {
    return objApp.LoadStringFromFile(filename, stringName);
}

/**
 * 导入指定路径中的文件
 * @param  {string}   path 指定路径
 */
function doImport(path) {
    // objApp.AppPath wiz的安装路径
    var templateFileName = objApp.AppPath + "templates\\new\\journal\\journal.template.htm";
    
    // 从一个text文件获得文字内容。可以有效地获得避免乱码。bstrFileName：文件名；返回值：text里面的文字
    var templateHtml = objCommon.LoadTextFromFile(templateFileName);
    
    var journalIniFileName = objApp.AppPath + "templates\\new\\journal.ini";

    ///枚举一个文件夹下面的文件。bstrPath：文件夹路径；bstrFileExt：扩展名，格式为*.txt;*.jpg;*.png或者*.*；vbIncludeSubFolders：是否包含子文件夹；
    ///返回值：枚举到的文件名。类型为一个安全数组，如果在javascript里面使用，请参阅本文后面部分。
    var objFiles = objCommon.EnumFiles(path, "*.html", false);
   
    //
    var arrFiles = objFiles;

    // 进度条
    var objProgress = objApp.CreateWizObject("WizKMControls.WizProgressWindow");
    objProgress.Show();
    objProgress.Max = arrFiles.length;
    objProgress.Title = "正在导入...";

    // 遍历文件
    for (var iFile = 0; iFile < arrFiles.length; iFile++) {
        var filename = arrFiles[iFile];

        // 进度条上显示当前处理的文件名
        objProgress.Text = filename;

        // 获得文件内容
        var html = objCommon.LoadTextFromFile(filename);

        //
        var objParts = objCommon.HtmlExtractTags(html, "div", "class", "qqshowbd");
        if (null == objParts)
            continue;

        var arrParts = objParts;
        for (var iPart = 0; iPart < arrParts.length; iPart++) {
            var part = arrParts[iPart];
            //
            var parthtml = part;
            //
            var title = getTitle(parthtml);
            //
            var detail = getDetail(parthtml);
            //
            var date = detail["date"];
            var category = detail["category"];
            var weather = detail["weather"];
            if (date == null)
                date = new Date();
            if (category == null)
                category = "未分类";
            if (weather == null)
                weather = "";
            //
            var date_string = DateToChineseStr(date);
            //
            var title = date_string + " " + title;
            //
            var location;
            if (asjournal) {
                location = "/My Journals/";
            } else {
                location = "/QQ记事本/" + category + "/";
            }
            location = location + date.getFullYear() + "-" + formatInt(date.getMonth() + 1) + "/";
            //
            var objFolder = objDatabase.GetFolderByLocation(location, true);
            //
            var objDoc = objFolder.CreateDocument2(title, "");
            if (asjournal) {
                objDoc.Type = "journal";
                setParamValue(objDoc, "journal-date", DateToStr(date) + " 00:00:00");
            }
            //
            var htmltext = "";
            if (asjournal) {
                htmltext = templateHtml;
                //
                htmltext = htmltext.replace("%title%", title);

                var resSunday = getLocalString(journalIniFileName, "resSunday");
                var resMonday = getLocalString(journalIniFileName, "resMonday");
                var resTuesday = getLocalString(journalIniFileName, "resTuesday");
                var resWednesday = getLocalString(journalIniFileName, "resWednesday");
                var resThursday = getLocalString(journalIniFileName, "resThursday");
                var resFriday = getLocalString(journalIniFileName, "resFriday");
                var resSaturday = getLocalString(journalIniFileName, "resSaturday");

                var weekarray = new Array(resSunday, resMonday, resTuesday, resWednesday, resThursday, resFriday, resSaturday);

                date_string = date_string + " " + weekarray[date.getDay()];

                //alert(date_string);

                htmltext = htmltext.replace("%date%", date_string);

                htmltext = htmltext.replace("%weather%", weather);
                htmltext = htmltext.replace("%text%", parthtml);
            } else {
                htmltext = parthtml;
            }

            // 更改文档标题和文件名
            objDoc.ChangeTitleAndFileName(title);

            /**
             * http://www.wiz.cn/manual/plugin/api/descriptions/IWizDocument.html
             * 
             * 更改文档数据，通过HTML文字内容来更新
             * [id(52), helpstring("method UpdateDocument3")] HRESULT UpdateDocument3([in] BSTR bstrHtml, [in] LONG nFlags);
             */
            objDoc.UpdateDocument3(htmltext, 0);
        }

        // 进度条进度移动到下一个为知
        objProgress.Pos = iFile + 1;
    }
}

/**
 * 点击对话框确认按钮
 */
function submitDialog() {
    var basepath = textFolder.value;

    // 如果没有选中路径，则返回
    if (basepath == null || basepath == "") {
        return;
    }

    // 转义路径
    if (basepath.charAt(basepath.length - 1) != '\\'){
        basepath = basepath + "\\";        
    }

    // 导入文件到选中的路径
    doImport(basepath);

    // 关闭对话框
    closeDialog();
}

/**
 * 关闭对话框
 */
function closeDialog() {
    objApp.Window.CloseHtmlDialog(document, ret);
}

/**
 * 点击之后，浏览文件夹
 */
function clickToBrowse() {
    /**
     * http://www.wiz.cn/manual/plugin/api/descriptions/IWizCommonUI.html
     * 
     * 提示用户选择一个磁盘文件夹。bstrDescription：描述；返回值：用户选择的文件夹
     * [id(16), helpstring("method SelectWindowsFolder")] HRESULT SelectWindowsFolder([in] BSTR bstrDescription, [out,retval] BSTR* pbtrFolderPath);
     */
    var importPath = objCommon.SelectWindowsFolder("选择文件夹");

    textFolder.value = importPath;
}
