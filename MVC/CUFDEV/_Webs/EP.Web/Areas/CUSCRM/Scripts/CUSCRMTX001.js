
var category = document.getElementById("Category");
var caseType = document.getElementById("CaseType");
var noPolicyNo = document.getElementById("NoPolicyNo");
var PolicyNo = document.getElementById("PolicyNo");
const redirectUrl = '/Account/Login';
var ownerMobile = "";
var PolicySet = new Set();
var Complainants = new Set();
var token = document.getElementById('__AjaxAntiForgeryForm').elements['__RequestVerificationToken'].value;
var stepFlag = "step1";
var CCCheckedComplainants = new Set();
var agentInfo;

history.pushState(null, document.title, location.href);
window.addEventListener('popstate', function (event) {
    history.pushState(null, document.title, location.href);
});

function GetRedirectUrl(url, keyword) {
    let point = url.indexOf(keyword);
    if (point != -1) {
        point += keyword.length;
        url = url.substring(0, point);
    }
    return url;
}

Date.prototype.addDays = function (days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

function convertTimestampToDateTime(timestamp) {
    if (timestamp != '' && timestamp != undefined && timestamp != null) {
        let date = new Date(parseInt(timestamp.substr(6)));
        let year = date.getFullYear();
        let month = addZeroPrefix(date.getMonth() + 1);
        let day = addZeroPrefix(date.getDate());
        let hours = addZeroPrefix(date.getHours());
        let minutes = addZeroPrefix(date.getMinutes());
        let seconds = addZeroPrefix(date.getSeconds());
        return year + "/" + month + "/" + day + " " + hours + ":" + minutes + ":" + seconds;
    }
    else {
        return '';
    }
}

function addZeroPrefix(number) {
    return number < 10 ? "0" + number : number;
}

function formatDate(date) {
    let d = new Date(date);
    let year = d.getFullYear();
    let month = addZeroPrefix(d.getMonth() + 1);
    let day = addZeroPrefix(d.getDate());
    return [year, month, day].join('/');
}

function getChainDate() {
    var now = new Date();
    var year = now.getFullYear() - 1911;
    var month = now.getMonth() + 1;
    var day = now.getDate();
    if (month.toString().length == 1) {
        month = '0' + month;
    }
    if (day.toString().length == 1) {
        day = '0' + day;
    }
    var dateTime = year + '/' + month + '/' + day;
    return dateTime;
}

function handleInput(event) {
    var input = event.target;
    var value = input.value;

    value = value.replace(/[^0-9\/]/g, '');

    if (value.length == 4) {
        value = value.slice(0, 3) + '/' + value.slice(3);
    }
    if (value.length == 7) {
        value = value.slice(0, 6) + '/' + value.slice(6);
    }

    input.value = value;
}

function handleKeyDown(event) {
    var key = event.key;

    if (key === 'Backspace') {
        var input = event.target;
        var selectionStart = input.selectionStart;
        var selectionEnd = input.selectionEnd;

        if (selectionStart !== selectionEnd) {
            var value = input.value;
            var newValue = value.slice(0, selectionStart) + value.slice(selectionEnd);
            input.value = newValue;
            input.selectionStart = input.selectionEnd = selectionStart;
            if (input.value.endsWith('/')) {
                selectionStart = input.selectionStart;
                input.value = newValue.slice(0, selectionStart - 1);
                input.selectionStart = input.selectionEnd = selectionStart - 1;
            }
            event.preventDefault();
        } else if (selectionStart > 0) {
            var value = input.value;
            var newValue = value.slice(0, selectionStart - 1) + value.slice(selectionStart);
            input.value = newValue;
            input.selectionStart = input.selectionEnd = selectionStart - 1;
            if (input.value.endsWith('/')) {
                selectionStart = input.selectionStart;
                input.value = newValue.slice(0, selectionStart - 1);
                input.selectionStart = input.selectionEnd = selectionStart - 1;
            }
            event.preventDefault();
        }
    } else if (key === '/') {
        event.preventDefault();
    }
}

function isValidDate(dateString) {
    if (dateString == undefined) {
        return false;
    }
    var date = new Date(dateString);

    if (isNaN(date.getTime())) {
        return false;
    }
    var dateObj = dateString.split('/');

    var limitInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    var theYear = parseInt(dateObj[0]);
    var theMonth = parseInt(dateObj[1]);
    var theDay = parseInt(dateObj[2]);
    var isLeap = new Date(theYear, 1, 29).getDate() === 29;

    if (isLeap) {
        limitInMonth[1] = 29;
    }

    return theDay <= limitInMonth[theMonth - 1];
}

function AddComplainants(ComplainantName) {
    if (!Complainants.has(ComplainantName)) {
        Complainants.add(ComplainantName);
    }
}
category.addEventListener('change', function () {
    postData('/CUSCRM/CUSCRMCommon/GetCaseType', { category: this.value}, (myJson) => {
        caseType.innerText = "";
        for (let i = 0; i < myJson.length; i++) {
            var item = myJson[i];
            var opt = document.createElement('option');
            opt.value = item.Value;
            opt.text = item.Text;
            caseType.appendChild(opt);
        }
    });
});

noPolicyNo.addEventListener('change', function () {
    if (this.checked) {
        PolicyNo.disabled = true;
        PolicyNo.value = "";
    } else {
        PolicyNo.disabled = false;
    }
});


document.getElementById('SelectAllPolicy').addEventListener('change', function () {
    let table = document.getElementById('ValidPolicyNoDatas')
    let rows = table.rows;
    for (let i = 1; i < rows.length; i++) {
        let chkObj = rows[i].cells[0].childNodes[0];
        if (chkObj != undefined && chkObj != null) {
            chkObj.checked = this.checked;
        }
    }
});

document.getElementById('btnQueryPolicy').addEventListener('click', function (e) {
    clearMessage();
    if (noPolicyNo.checked) {
        document.getElementsByName("step1").forEach((e) => { e.classList.add("none"); });
        document.getElementsByName("step1-1").forEach((e) => { e.classList.add("none"); });
        document.getElementsByName("step2").forEach((e) => { e.classList.add("none"); });
        document.getElementsByName("step2-3").forEach((e) => { e.classList.add("none"); });
        document.getElementsByName("step2-4").forEach((e) => { e.classList.remove("none"); });
    }
    else {
        if (PolicyNo.value.trim() == "") {
            appendMessage('請輸入保單號碼');
            return;
        }

        CheckPolicyNoNotClosed();
    }
});

function GetPolicyDatasByOwnerID(ownerID) {
    document.getElementsByName("step2-1").forEach((e) => { e.classList.add("none"); });

    postData('/CUSCRM/CUSCRMTX001/GetPolicyDatas', { ownerID: ownerID }, (policyJson) => {

        let tbodyValid = document.getElementById('ValidPolicyNoDatas');
        BindPolicyDatas(tbodyValid, policyJson.ValidPolicys);

        let tbodyInvalid = document.getElementById('InvalidPolicyNoDatas');
        BindPolicyDatas(tbodyInvalid, policyJson.InvalidPolicys);

        BindSelectPolicy();

    }, null, () => { document.getElementsByName("step2-2").forEach((e) => { e.classList.remove("none"); });});
}

function BindPolicyDatas(table, data) {
    tbody = table.getElementsByTagName('tbody')[0]
    tbody.innerText = "";
    for (let i = 0; i < data.length; i++) {
        let item = data[i];
        let newRow = tbody.insertRow();
        let newCheckBox = document.createElement("input");
        let itemValue = JSON.stringify(item);
        newCheckBox.setAttribute("type", "checkbox");
        newCheckBox.setAttribute("name", "CheckedPolicy");
        newCheckBox.className = "form-check-input";
        newCheckBox.setAttribute("data-value", itemValue);
        if (PolicySet.has(itemValue)) {
            newCheckBox.checked = true;
            PolicySet.delete(itemValue);
        }
        else if (PolicyNo.value.trim().toUpperCase() === item.PolicyNo.trim().toUpperCase()) {
            newCheckBox.checked = true;
        }

        let newCell = newRow.insertCell();
        newCell.appendChild(newCheckBox);

        newCell = newRow.insertCell();
        let newSerialText = document.createTextNode(i + 1);
        newCell.appendChild(newSerialText);

        newCell = newRow.insertCell();
        let newPolicyNoText = document.createTextNode(StringNullToEmpty(item.PolicyNo));
        newCell.appendChild(newPolicyNoText);

        newCell = newRow.insertCell();
        let newCompanyNameText = document.createTextNode(StringNullToEmpty(item.CompanyName));
        newCell.appendChild(newCompanyNameText);

        newCell = newRow.insertCell();
        let newOwnerText = document.createTextNode(StringNullToEmpty(item.Owner));
        newCell.appendChild(newOwnerText);

        newCell = newRow.insertCell();
        let newInsuredText = document.createTextNode(StringNullToEmpty(item.Insured));
        newCell.appendChild(newInsuredText);

        newCell = newRow.insertCell();
        let newIssDateText = document.createTextNode(StringNullToEmpty(item.IssDate));
        newCell.appendChild(newIssDateText);

        newCell = newRow.insertCell();
        let newStatusNameText = document.createTextNode(StringNullToEmpty(item.StatusName));
        newCell.appendChild(newStatusNameText);

        newCell = newRow.insertCell();
        let newModPremText = document.createTextNode(item.ModPrem.toLocaleString('en-US'));
        newCell.appendChild(newModPremText);

        newCell = newRow.insertCell();
        let newModxNameText = document.createTextNode(StringNullToEmpty(item.ModxName));
        newCell.appendChild(newModxNameText);

        let wcName = StringNullToEmpty(item.SUWCName);
        if (wcName == '') {
            wcName = StringNullToEmpty(item.WCName);
        }

        newCell = newRow.insertCell();
        let newWCNameText = document.createTextNode(wcName);
        newCell.appendChild(newWCNameText);

        newCell = newRow.insertCell();
        let newAgentNameText = document.createTextNode(StringNullToEmpty(item.AgentName));
        newCell.appendChild(newAgentNameText);

        newCell = newRow.insertCell();
        let newSUAgentNameText = document.createTextNode(StringNullToEmpty(item.SUAgentName));
        newCell.appendChild(newSUAgentNameText);
    }

    if (data.length == 0) {
        let newRow = tbody.insertRow();
        let newCell = newRow.insertCell();
        newCell.colSpan = 14;
        let newEmptyText = document.createTextNode("無符合資料");
        newCell.appendChild(newEmptyText);
    }
    else {
        RowClick(table);
    }
}

document.getElementById('btnQueryOwner').addEventListener('click', function (e) {
    clearMessage();
    let checkedPolicy = document.querySelector('input[name="CheckedOwner"]:checked');
    if (checkedPolicy == undefined || checkedPolicy == null) {
        appendMessage('請勾選一筆資料');
        return;
    }
    let ownerID = document.querySelector('input[name="CheckedOwner"]:checked').value;
    if (ownerID == "" || ownerID == undefined || ownerID == null) {
        appendMessage('請勾選一筆資料');
        return;
    }
    else {
        GetPolicyDatasByOwnerID(ownerID);
        document.getElementById('DuplicatePolicyNo').getElementsByTagName('tbody')[0].innerHTML = "";
    }
});

document.getElementById('btnAddPolicy').addEventListener('click', function (e) {
    clearMessage();
    if (document.querySelectorAll('input[name="CheckedPolicy"]:checked').length == 0) {
        appendMessage("至少請勾選一筆保單資料");
        return;
    }
    else {
        document.querySelectorAll('input[name="CheckedPolicy"]:checked').forEach((e) => {
            let dataValue = e.getAttribute('data-value');
            PolicySet.add(dataValue);
        });
        PolicyNo.value = "";
        noPolicyNo.checked = false;
        document.getElementsByName("step2-2").forEach((e) => { e.classList.add("none"); });
        document.getElementsByName("step2-3").forEach((e) => { e.classList.remove("none"); });
        document.getElementById('ValidPolicyNoDatas').getElementsByTagName('tbody')[0].innerHTML = "";
        document.getElementById('InvalidPolicyNoDatas').getElementsByTagName('tbody')[0].innerHTML = "";
    }
});

document.getElementById('btnCaseContent').addEventListener('click', function (e) {
    clearMessage();
    if (document.querySelectorAll('input[name="CheckedPolicy"]:checked').length == 0) {
        appendMessage("至少請勾選一筆保單資料");
        return;
    }
    else {
        document.querySelectorAll('input[name="CheckedPolicy"]:checked').forEach((e) => {
            let dataValue = e.getAttribute('data-value');
            PolicySet.add(dataValue)
        });
        document.getElementsByName("step2-2").forEach((e) => { e.classList.add("none"); });
        stepFlag = "step2-2";
    }
    ShowTypeContent();
});

document.getElementById('btnNoDataNext').addEventListener('click', function (e) {
    clearMessage();
    let issDate = document.getElementById('IssDate').value.trim();

    if (document.getElementById('OwnerName').value.trim() == "") {
        appendMessage('要保人姓名／申訴人不能為空白');
        return;
    }

    if (issDate != "") {
        let dateArray = issDate.split('/');
        if (dateArray.length != 3) {
            appendMessage('生效日格式有誤');
            return;
        }
        else {
            if (!isValidDate((parseInt(dateArray[0]) + 1911).toString() + '/' + dateArray[1] + '/' + dateArray[2])) {
                appendMessage('生效日格式有誤');
                return;
            }
            else {
                issDate = issDate.padStart(10, '0');
            }
        }
    }

    if (document.getElementById('AgentCode').value.trim() == "") {
        appendMessage('業務人員代碼不能為空白');
        return;
    }

    let agentCode = document.getElementById('AgentCode').value.trim();

    postData('/CUSCRM/CUSCRMTX001/GetAgentInfo', { agentCode: agentCode }, (data) => {
        if (data.result == true) {
            agentInfo = data.data;
            CheckPolicyNoNotClosedYouself();
        }
        else {
            appendMessage('業務員代碼不存在，請重新輸入!');
            return;
        }
    });
    

});

[document.getElementById('btnCancelQueryOwner'), document.getElementById('btnBackQueryPanel')].map(elm => {
    elm.addEventListener('click',  () => {
        Cancel(true);
    });
});

[document.getElementById('btnCSCancel'), document.getElementById('btnSSCancel'), document.getElementById('btnCCCancel'), document.getElementById('btnSCCancel')].map(elm => {
    elm.addEventListener('click', () => {
        document.getElementById("CSContent").classList.add("none");
        document.getElementById("SSContent").classList.add("none");
        document.getElementById("CCContent").classList.add("none");
        document.getElementById("SCContent").classList.add("none");

        if (caseType.value == "CC") {
            CCCheckedComplainants.clear();
            document.querySelectorAll('input[name="CheckedComplainants"]:checked').forEach((e) => { CCCheckedComplainants.add(e.value); });
        }

        if (stepFlag == "step2-2") {

            RemoveExistsPolicyDatas('ValidPolicyNoDatas');
            RemoveExistsPolicyDatas('InvalidPolicyNoDatas');

            BindSelectPolicy();

            document.getElementsByName("step2-2").forEach((e) => { e.classList.remove("none"); });
        }
        else {
            Complainants.clear();
            document.getElementsByName("step2-4").forEach((e) => { e.classList.remove("none"); });
        }
    });
});

document.getElementById('btnCSNoNotify').addEventListener('click', () => {
    SaveCSContent(3, (CaseNo) => {
        appendMessage('受理序號「' + CaseNo + '」新增成功');
        Cancel(false);
    });
});

document.getElementById('btnCSNotify').addEventListener('click', () => {
    SaveCSContent(0, (CaseNo) => {
        notify('受理序號「' + CaseNo + '」新增成功', () => {
            useBlockUI();
            window.location.href = 'CUSCRMTX002/Notify?No=' + CaseNo
        });
    });
});

document.getElementById('btnSSNotify').addEventListener('click', () => {
    SaveSSContent(3, (CaseNo) => {
        appendMessage('受理序號「' + CaseNo + '」新增成功');
        Cancel(false);
    });
});

document.getElementById('btnCCNoNotify').addEventListener('click', () => {
    SaveCCContent(3, (CaseNo) => {
        appendMessage('受理序號「' + CaseNo + '」新增成功',true);
        Cancel(false);
        window.location.href = 'CUSCRMTX002/CRMEAppeal?No=' + CaseNo 
    });
});

document.getElementById('btnCCNotify').addEventListener('click', () => {
    SaveCCContent(0, (CaseNo) => {
        notify('受理序號「' + CaseNo + '」新增成功', () => {
            useBlockUI();
            //window.location.href = 'CUSCRMTX002/Notify?No=' + CaseNo
            window.location.href = 'CUSCRMTX002/CRMEAppeal?No=' + CaseNo
        });
    });
});

document.getElementById('btnSCNoNotify').addEventListener('click', () => {
    SaveSCContent(3, true, (CaseNo) => {
        appendMessage('受理序號「' + CaseNo + '」新增成功');
        Cancel(false);
    });
});

document.getElementById('btnSCNotify').addEventListener('click', () => {
    SaveSCContent(0, false, (CaseNo) => {
        notify('受理序號「' + CaseNo + '」新增成功', () => {
            useBlockUI();
            window.location.href = 'CUSCRMTX002/Notify?No=' + CaseNo
        });
    });
});

function StringNullToEmpty(value) {
    if (value == null || value == undefined) {
        value = "";
    }
    return value;
}

document.querySelectorAll("input[name='CSSourceID']").forEach((input) => {
    input.addEventListener('change', (e) => {
        if (e.target.getAttribute('data-name') == '保公CSR') {
            document.getElementById('CSCompanyPanel').classList.remove("none");
        }
        else {
            document.getElementById('CSCompanyPanel').classList.add("none");
        }

        sourceValue = e.target.getAttribute('data-name');

        let contentCaseType = document.querySelector('input[name="CSCaseTypeID"]:checked');
        let contentCaseTypeValue;
        if (contentCaseType != null && contentCaseType != undefined) {
            contentCaseTypeValue = contentCaseType.getAttribute('data-name')
        }

        let caller = document.querySelector('input[name="CSCallerID"]:checked');
        let callerValue;
        if (caller != null && caller != undefined) {
            callerValue = caller.getAttribute('data-name')
        }
        SetCSDefaultContent(sourceValue, callerValue, contentCaseTypeValue);
    });
});

document.querySelectorAll("input[name='CSCallerID']").forEach((input) => {
    input.addEventListener('change', (e) => {
        let source = document.querySelector('input[name="CSSourceID"]:checked');
        let sourceValue;
        if (source != null && source != undefined) {
            sourceValue = source.getAttribute('data-name')
        }
        let contentCaseType = document.querySelector('input[name="CSCaseTypeID"]:checked');
        let contentCaseTypeValue;
        if (contentCaseType != null && contentCaseType != undefined) {
            contentCaseTypeValue = contentCaseType.getAttribute('data-name')
        }
        let callerValue = e.target.getAttribute('data-name');

        SetCSDefaultContent(sourceValue, callerValue, contentCaseTypeValue);
        
    });
});

document.querySelectorAll("input[name='CSCaseTypeID']").forEach((input) => {
    input.addEventListener('change', (e) => {
        let source = document.querySelector('input[name="CSSourceID"]:checked');
        let sourceValue;
        if (source != null && source != undefined) {
            sourceValue = source.getAttribute('data-name')
        }

        let contentCaseTypeValue = e.target.getAttribute('data-name');

        let caller = document.querySelector('input[name="CSCallerID"]:checked');
        let callerValue;
        if (caller != null && caller != undefined) {
            callerValue = caller.getAttribute('data-name')
        }
        SetCSDefaultContent(sourceValue, callerValue, contentCaseTypeValue);
    });
});

document.getElementById("CSCompany").addEventListener('change', (e) => {
        let source = document.querySelector('input[name="CSSourceID"]:checked');
        let sourceValue;
        if (source != null && source != undefined) {
            sourceValue = source.getAttribute('data-name')
        }

        let contentCaseType = document.querySelector('input[name="CSCaseTypeID"]:checked');
        let contentCaseTypeValue;
        if (contentCaseType != null && contentCaseType != undefined) {
            contentCaseTypeValue = contentCaseType.getAttribute('data-name')
        }

        let caller = document.querySelector('input[name="CSCallerID"]:checked');
        let callerValue;
        if (caller != null && caller != undefined) {
            callerValue = caller.getAttribute('data-name')
        }
        SetCSDefaultContent(sourceValue, callerValue, contentCaseTypeValue);
});


document.querySelectorAll("input[name='CCSourceID']").forEach((input) => {
    input.addEventListener('change', (e) => {

        sourceValue = e.target.getAttribute('data-name');
        let content = document.getElementById('CCSourceDesc');

        document.getElementById('CCSourceDescDiv').classList.add('none');
        switch (sourceValue) {
            case '申訴函':
                content.value = '保戶申訴函';
                break;
            case '保戶來電':
                content.value = '保戶來電申訴';
                break;
            case '保險局':
                content.value = '保險局轉知保戶申訴資料';
                break;
            case '評議中心':
                content.value = '評議中心轉知保戶申訴資料';
                break;
            case '存證信函':
                content.value = '保戶存證信函';
                break;
            case '律師函':
                content.value = '保戶律師函';
                break;
            case '官網':
                content.value = '保戶於官網留言申訴';
                break;
            case '其他':
                content.value = '';
                document.getElementById('CCSourceDescDiv').classList.remove('none');
            default:
                break;
        }
    });
});

document.querySelectorAll("input[name='SCSourceID']").forEach((input) => {
    input.addEventListener('change', (e) => {

        sourceValue = e.target.getAttribute('data-name');
        let content = document.getElementById('SCSourceDesc');

        document.getElementById('SCSourceDescDiv').classList.add('none');
        switch (sourceValue) {
            case '申訴函':
                content.value = '申訴函';
                break;
            case '業連申訴':
                content.value = '業連申訴';
                break;
            case '其他':
                content.value = '';
                document.getElementById('SCSourceDescDiv').classList.remove('none');
            default:
                break;
        }
    });
});

function RowClick(table ) {
    let rows = table.rows;
    for (let i = 1; i < rows.length; i++) {

        let cells = rows[i].cells;
        for (let j = 1; j < cells.length; j++) {
            cells[j].onclick = (function () {
                return function () {
                    let chkObj = rows[i].cells[0].childNodes[0];
                    if (chkObj.checked == true) {
                        chkObj.checked = false;
                    }
                    else {
                        chkObj.checked = true;
                    }
                }
            })(i);
        }
    }
}

function NotifyRowClick(table) {
    let rows = table.rows;
    for (let i = 1; i < rows.length; i++) {

        let cells = rows[i].cells;
        for (let j = 1; j < cells.length; j++) {
            cells[j].onclick = (function () {
                return function () {
                    let caseNo = rows[i].cells[2].childNodes[0];
                    useBlockUI();
                    window.location.href = 'CUSCRMTX002/Notify?No=' + caseNo.nodeValue;
                }
            })(i);
        }
    }
}

function CCNotifyRowClick(table) {
    let rows = table.rows;
    for (let i = 1; i < rows.length; i++) {

        let cells = rows[i].cells;
        for (let j = 1; j < cells.length; j++) {
            cells[j].onclick = (function () {
                return function () {
                    let caseNo = rows[i].cells[2].childNodes[0];
                    useBlockUI();
                    window.location.href = 'CUSCRMTX002/CRMEAppeal?No=' + caseNo.nodeValue;
                }
            })(i);
        }
    }
}

function SetCSDefaultContent(sourceValue, callerValue, contentCaseTypeValue) {
    let caseA = ['一般', '減額繳清', '契變', '保費', '自動墊繳', '理賠', '復效', '其他'];
    let caseB = ['解約/部解'];
    let caseC = ['更換業代'];
    let caseD = ['理賠-身故'];
    let currentDate = getChainDate();
    let fromName;
    switch (sourceValue) {
        case "保公CSR":
            let company = document.getElementById('CSCompany').value;
            let companyValue = document.querySelector('#CSCompany option[value="' + company + '"]').innerText;
            fromName = '接獲' + companyValue + '通知';
            break;
        case "0800":
            fromName = '接獲' + callerValue + '來電';
            break;
        case "官網":
        case "Line":
            fromName = '接獲' + callerValue + '於' + sourceValue + '留言';
            break;
        case "總公司電訪":
            fromName = '本部執行電訪';
            break;
        case '其他':
            fromName = '';
            break;
        default:
            break;
    }

    let content = document.getElementById('CSCaseContent');

    if (caseA.includes(contentCaseTypeValue)) {

        content.value = '一、' + currentDate + fromName + '\n' + '二、保戶手機：' + ownerMobile + '\n' + '三、請　台端儘速與保戶聯絡，並於回覆欄告知以下2點，以利結案。謝謝!' + '\n' + '1、與保戶聯繫日期、時間' + '\n' + '2、請詳述本案處理經過與結果';
    }
    else if (caseB.includes(contentCaseTypeValue)) {
        content.value = '一、' + currentDate + fromName + '\n' + '二、保戶手機：' + ownerMobile + '\n' + '三、請　台端儘速與保戶聯絡進行保全作業，並於回覆欄告知以下2點，以利結案。謝謝!' + '\n' + '1、與保戶聯繫日期、時間' + '\n' + '2、請詳述本案處理經過與結果';
    }
    else if (caseC.includes(contentCaseTypeValue)) {
        content.value = '一、' + currentDate + fromName + '\n' + '二、因保戶訴求變更上述保單接續服務人員，請於回覆欄告知後續單位處理結果，以利結案。謝謝!' + '\n' + '1、如同意保戶變更服務人員， 請於回覆欄告知「同意保戶訴求」。' + '\n' + '2、如不同意保戶變更，請回覆E化問卷完成日期（客服部將直接於系統查詢），或後續將如何服務保戶。';
    }
    else if (caseD.includes(contentCaseTypeValue)) {
        content.value = '一、本部受理' + fromName + '：內政部通報被保險人已身故，但尚未收到理賠文件\n' + '二、請協助於回覆欄回覆以下2點，以利結案。謝謝' + '\n' + '1、是否已協助保戶家屬辦理理賠事宜，如有，請回覆送件日' + '\n' + '2、如否，請回覆說明未辦理理賠之原因。';
    }
}

function Cancel(needClrMsg) {

    if (needClrMsg) {
        clearMessage();
    }
    tempMobile = "";
    ownerMobile = "";
    PolicySet.clear();
    Complainants.clear();
    CCCheckedComplainants.clear();
    agentInfo = null;
    document.getElementById('ValidPolicyNoDatas').getElementsByTagName('tbody')[0].innerHTML = "";
    document.getElementById('InvalidPolicyNoDatas').getElementsByTagName('tbody')[0].innerHTML = "";
    document.getElementById('InvalidPolicyNoDatas').getElementsByTagName('tbody')[0].innerHTML = "";
    document.getElementById('DuplicatePolicyNo').getElementsByTagName('tbody')[0].innerHTML = "";
    document.getElementById('SelectPolicyNoDatas').getElementsByTagName('tbody')[0].innerHTML = "";
    document.getElementById('OwnerName').value = "";
    document.getElementById('InsuredName').value = "";
    document.getElementById('IssDate').value = "";
    document.getElementById('PolicyNo2').value = "";
    document.getElementById('AgentCode').value = "";

    category.value = 0;
    noPolicyNo.checked = false;
    var event = new Event('change');
    category.dispatchEvent(event);
    noPolicyNo.dispatchEvent(event);
    PolicyNo.value = "";

    document.getElementsByName("step2-1").forEach((e) => { e.classList.add("none"); });
    document.getElementsByName("step2-2").forEach((e) => { e.classList.add("none"); });
    document.getElementsByName("step2-4").forEach((e) => { e.classList.add("none"); });
    document.getElementById("CSContent").classList.add("none");
    document.getElementById("SSContent").classList.add("none");
    document.getElementById("CCContent").classList.add("none");
    document.getElementById("SCContent").classList.add("none");

    ClearCSContent();
    ClearSSContent();
    ClearCCContent();
    ClearSCContent();

    document.getElementsByName("step1").forEach((e) => { e.classList.remove("none"); });
    document.getElementsByName("step2").forEach((e) => { e.classList.remove("none"); });
    document.getElementsByName("step2-3").forEach((e) => { e.classList.remove("none"); });
    GetWaitNofityDatas();
    GetCCWaitNofityDatas();
}

function ShowTypeContent() {

    let caseType = document.getElementById('CaseType').value;
    if (caseType == 'CS') {
        let [first] = PolicySet;
        ownerMobile = JSON.parse(first).OwnerMobile;
        document.getElementById("CSContent").classList.remove("none");
    }
    else if (caseType == 'SS') {
        document.getElementById("SSContent").classList.remove("none");
    }
    else if (caseType == 'CC') {
        PolicyDatasGetComplainants();
        let comliants = document.getElementById('CCComplainantsDiv');
        comliants.innerHTML = '';
        let index = 0;
        Array.from(Complainants).forEach((e) => {
            if (e != '' && e != undefined && e != null) {
                let compliantDiv = document.createElement("div");
                compliantDiv.classList.add('form-check');
                compliantDiv.classList.add('form-check-inline');
                let compliantCheckBox = document.createElement("input");
                compliantCheckBox.setAttribute("type", "checkbox");
                compliantCheckBox.setAttribute("name", "CheckedComplainants");
                compliantCheckBox.setAttribute("id", "CheckedComplainants" + index.toString());
                compliantCheckBox.className = "form-check-input";
                compliantCheckBox.value = e;
                if (CCCheckedComplainants.has(e)) {
                    compliantCheckBox.checked = true;
                }
                let compliantLabel = document.createElement("label");
                compliantLabel.classList.add('form-check-label');
                compliantLabel.setAttribute('for', "CheckedComplainants" + index.toString());
                compliantLabel.textContent = e;
                compliantDiv.appendChild(compliantCheckBox);
                compliantDiv.appendChild(compliantLabel);
                comliants.appendChild(compliantDiv);
                index++;
            }
        });
        document.getElementById("CCContent").classList.remove("none");;
    }
    else if (caseType == 'SC') {
        PolicyDatasGetComplainants();
        document.getElementById("SCContent").classList.remove("none");;
    }
}

function ClearCSContent() {
    document.querySelectorAll('input[name="CSSourceID"]').forEach((e) => { e.checked = false; });
    document.querySelectorAll('input[name="CSCaseTypeID"]').forEach((e) => { e.checked = false; });
    document.querySelectorAll('input[name="CSCallerID"]')[0].checked = true;
    document.getElementById('CSCompany').selectedIndex = 0;
    document.getElementById('CSCaseContent').value = '';
    document.getElementById('CSCompanyPanel').classList.add("none");
}

function ClearSSContent() {
    document.querySelectorAll('input[name="SSSourceID"]')[0].checked = true;
    document.querySelectorAll('input[name="SSCaseTypeID"]').forEach((e) => { e.checked = false; });
    document.getElementById('SSCaseContent').value = '';
}

function ClearCCContent() {
    document.querySelectorAll('input[name="CCSourceID"]').forEach((e) => { e.checked = false; });
    document.querySelectorAll('input[name="CCCaseTypeID"]').forEach((e) => { e.checked = false; });
    document.querySelectorAll('input[name="CCCaseCategoryID"]').forEach((e) => { e.checked = false; });
    document.getElementById('CCReceiveDateTime').value = '';
    document.getElementById('CCReplayDDLDateTime').value = '';
    document.getElementById('CCSourceDescDiv').classList.add('none');
    document.getElementById('CCSourceDesc').value = '';
    document.getElementById('SCSourceDescDiv').classList.add('none');
    document.getElementById('SCSourceDesc').value = '';


    document.getElementById('CCBizContract').checked = false;
    document.getElementById('CCSolicitingRpt').checked = false;
    document.getElementById('CCComplainantsDiv').innerHTML = '';
    document.getElementById('CCCaseContent').value = '';
}

function ClearSCContent() {
    document.querySelectorAll('input[name="SCSourceID"]').forEach((e) => { e.checked = false; });
    document.querySelectorAll('input[name="SCCaseTypeID"]').forEach((e) => { e.checked = false; });
    document.getElementById('SCReceiveDateTime').value = '';

    document.getElementById('SCBizContract').checked = false;
    document.getElementById('SCSolicitingRpt').checked = false;
    document.getElementById('SCCaseContent').value = '';
}

function SaveCSContent(status, callbackFunc) {
    clearMessage();
    if (document.querySelector('input[name="CSSourceID"]:checked') == undefined) {
        appendMessage('請選擇資料來源!');
        return;
    }

    if (document.querySelector('input[name="CSCaseTypeID"]:checked') == undefined) {
        appendMessage('請選擇案件類別!');
        return;
    }

    if (document.getElementById('CSCaseContent').value.trim() == '') {
        appendMessage('照會內容不能為空白!');
        return;
    }

    let sourceID = document.querySelector('input[name="CSSourceID"]:checked').value;
    let companyCode = "";
    if (document.querySelector('input[name="CSSourceID"]:checked').getAttribute('data-name') == '保公CSR') {
        companyCode = document.getElementById('CSCompany').value;
    }
    let callerID = document.querySelector('input[name="CSCallerID"]:checked').value;
    let caseTypeID = document.querySelector('input[name="CSCaseTypeID"]:checked').value;
    let content = document.getElementById('CSCaseContent').value;
    let policies = PolicyDataToJsonArray();

    postData('/CUSCRM/CUSCRMTX001/CreateCase', { Type: caseType.value, SourceID: sourceID, CompanyCode: companyCode, CallerID: callerID, CaseTypeID: caseTypeID, Content: content, Status: status, policies: JSON.stringify(policies) }, (myJson) => {
        if (callbackFunc != null) {
            callbackFunc(myJson.No);
        }
    });
}


function SaveSSContent(status, callbackFunc) {

    clearMessage();
    if (document.querySelector('input[name="SSSourceID"]:checked') == undefined) {
        appendMessage('請選擇資料來源!');
        return;
    }

    if (document.querySelector('input[name="SSCaseTypeID"]:checked') == undefined) {
        appendMessage('請選擇案件類別!');
        return;
    }

    if (document.getElementById('SSCaseContent').value.trim() == '') {
        appendMessage('案件內容不能為空白!');
        return;
    }

    let sourceID = document.querySelector('input[name="SSSourceID"]:checked').value;
    let caseTypeID = document.querySelector('input[name="SSCaseTypeID"]:checked').value;
    let content = document.getElementById('SSCaseContent').value;
    let policies = PolicyDataToJsonArray();

    postData('/CUSCRM/CUSCRMTX001/CreateCase', { Type: caseType.value, SourceID: sourceID, CaseTypeID: caseTypeID, Content: content, Status: status, policies: JSON.stringify(policies) }, (myJson) => {
        if (callbackFunc != null) {
            callbackFunc(myJson.No);
        }
    });
}


function SaveCCContent(status, callbackFunc) {

    clearMessage();

    let Complainant = '';

    if (document.getElementById('CCReceiveDateTime').value.trim() == "") {
        appendMessage('移交日期不為空白!');
        return;
    }
    else {
        if (!isValidDate(document.getElementById('CCReceiveDateTime').value)) {
            appendMessage('移交日期格式有誤!');
            return;
        }
    }

    if (document.getElementById('CCReplayDDLDateTime').value.trim() == "") {
        appendMessage('回文期限不為空白!');
        return;
    }
    else {
        if (!isValidDate(document.getElementById('CCReplayDDLDateTime').value)) {
            appendMessage('回文期限格式有誤!');
            return;
        }
    }

    if (document.querySelector('input[name="CCSourceID"]:checked') == undefined) {
        appendMessage('請選擇資料來源!');
        return;
    }

    if (document.getElementById('CCSourceDesc').value.trim() == '') {
        appendMessage('資料來源為其他時，其他的來源內容不能為空白!');
        return;
    }

    if (document.querySelector('input[name="CCCaseCategoryID"]:checked') == undefined) {
        appendMessage('請選擇案件類型!');
        return;
    }

    if (document.querySelector('input[name="CCCaseTypeID"]:checked') == undefined) {
        appendMessage('請選擇案件類別!');
        return;
    }

    if (document.querySelectorAll('input[name="CheckedComplainants"]:checked').length == 0) {
        appendMessage('請選擇申訴人!');
        return;
    }
    else {
        let ccComplainantList = [];
        document.querySelectorAll('input[name="CheckedComplainants"]:checked').forEach((e) => { ccComplainantList.push(e.value); });
        Complainant = ccComplainantList.join('、');
    }

    if (document.getElementById('CCCaseContent').value.trim() == '') {
        appendMessage('案件內容不能為空白!');
        return;
    }

    let bizContract = GetCheckedYesValue('CCBizContract');
    let solicitingRpt = GetCheckedYesValue('CCSolicitingRpt');

    if (bizContract == '' && solicitingRpt == '') {
        appendMessage('回覆方式至少需勾選一種!');
        return;
    }

    let sourceID = document.querySelector('input[name="CCSourceID"]:checked').value;
    let categoryID = document.querySelector('input[name="CCCaseCategoryID"]:checked').value;
    let caseTypeID = document.querySelector('input[name="CCCaseTypeID"]:checked').value;
    let receiveDateTime = document.getElementById('CCReceiveDateTime').value;
    let replayDDLDateTime = document.getElementById('CCReplayDDLDateTime').value;
    let content = document.getElementById('CCCaseContent').value;
    let sourceDesc = document.getElementById('CCSourceDesc').value;
    
    let policies = PolicyDataToJsonArray();

    postData('/CUSCRM/CUSCRMTX001/CreateCase', { Type: caseType.value, ReceiveDateTime: receiveDateTime, ReplayDDLDateTime: replayDDLDateTime, SourceID: sourceID, SourceDesc: sourceDesc, CaseTypeID: caseTypeID, CaseCategoryID: categoryID, Content: content, Status: status, BizContract: bizContract, SolicitingRpt: solicitingRpt, Complainant: Complainant, policies: JSON.stringify(policies) }, (myJson) => {
        if (callbackFunc != null) {
            callbackFunc(myJson.No);
        }
    });
}


function SaveSCContent(status, needConfirm, callbackFunc) {

    clearMessage();

    if (document.getElementById('SCReceiveDateTime').value.trim() == "") {
        appendMessage('收文日期不為空白!');
        return;
    }
    else {
        if (!isValidDate(document.getElementById('SCReceiveDateTime').value)) {
            appendMessage('收文日期格式有誤!');
            return;
        }
    }

    if (document.querySelector('input[name="SCSourceID"]:checked') == undefined) {
        appendMessage('請選擇資料來源!');
        return;
    }

    if (document.getElementById('SCSourceDesc').value.trim() == '') {
        appendMessage('資料來源為其他時，其他的來源內容不能為空白!');
        return;
    }

    if (document.querySelector('input[name="SCCaseTypeID"]:checked') == undefined) {
        appendMessage('請選擇案件類別!');
        return;
    }


    if (document.getElementById('SCCaseContent').value.trim() == '') {
        appendMessage('案件內容不能為空白!');
        return;
    }

    let bizContract = GetCheckedYesValue('SCBizContract');
    let solicitingRpt = GetCheckedYesValue('SCSolicitingRpt');

    if (bizContract == '' && solicitingRpt == '') {
        appendMessage('回覆方式至少需勾選一種!');
        return;
    }

    if (needConfirm) {
        let text = '確定不通知業務員?'
        confirm(text, (confirmed) => {
            if (confirmed) {
                SaveSCProcess(status, callbackFunc);
            }
        })
    }
    else {
        SaveSCProcess(status, callbackFunc);
    }
}

function SaveSCProcess(status , callbackFunc) {
    let sourceID = document.querySelector('input[name="SCSourceID"]:checked').value;
    let caseTypeID = document.querySelector('input[name="SCCaseTypeID"]:checked').value;
    let receiveDateTime = document.getElementById('SCReceiveDateTime').value;
    let content = document.getElementById('SCCaseContent').value;
    let bizContract = GetCheckedYesValue('SCBizContract');
    let solicitingRpt = GetCheckedYesValue('SCSolicitingRpt');
    let sourceDesc = document.getElementById('SCSourceDesc').value;

    let policies = PolicyDataToJsonArray();

    postData('/CUSCRM/CUSCRMTX001/CreateCase', { Type: caseType.value, ReceiveDateTime: receiveDateTime, SourceID: sourceID, SourceDesc :sourceDesc, CaseTypeID: caseTypeID, Content: content, Status: status, BizContract: bizContract, SolicitingRpt: solicitingRpt, Complainant: Complainants.values().next().value, policies: JSON.stringify(policies) }, (myJson) => {
        if (callbackFunc != null) {
            callbackFunc(myJson.No);
        }
    });
}

$('#CCReceiveDateTime').on('change', (e) => {
    if (isValidDate(e.target.value)) {
        document.getElementById('CCReplayDDLDateTime').value = formatDate(new Date(e.target.value).addDays(29));
    }
});

const confirm = (text, callback, yesName, noName) => {
    let confirm = document.querySelector('.confirm')
    let backdrop = document.querySelector('.backdrop')
    let confirmaffirmative = document.querySelector('.confirm-affirmative')
    let confirmnegative = document.querySelector('.confirm-negative')

    if (yesName == null || yesName == undefined) {
        confirmaffirmative.innerHTML = '是';
    }
    else {
        confirmaffirmative.innerHTML = yesName;
    }
    if (noName == null || noName == undefined) {
        confirmnegative.innerHTML = '否';
    }
    else {
        confirmnegative.innerHTML = noName;
    }
    confirm.classList.remove('none')
    backdrop.classList.remove('none')
    confirmaffirmative.classList.remove('none')
    confirmnegative.classList.remove('none')

    let confirmContent = confirm.querySelector('.text-content')
    confirmContent.textContent = text

    confirm.addEventListener('click', (e) => {
        let target = e.target

        if (target.className === 'confirm-affirmative') {
            confirm.classList.add('none')
            backdrop.classList.add('none')
            confirmaffirmative.classList.add('none')
            confirmnegative.classList.add('none')
            callback(true)
        }

        if (target.className === 'confirm-negative') {
            confirm.classList.add('none')
            backdrop.classList.add('none')
            confirmaffirmative.classList.add('none')
            confirmnegative.classList.add('none')
            callback(false)
        }
    })
}

const notify = (text, callback) => {
    let confirm = document.querySelector('.confirm')
    let backdrop = document.querySelector('.backdrop')
    let gonofity = document.querySelector('.go-notify')

    confirm.classList.remove('none')
    backdrop.classList.remove('none')
    gonofity.classList.remove('none')

    let confirmContent = confirm.querySelector('.text-content')
    confirmContent.textContent = text

    confirm.addEventListener('click', (e) => {
        let target = e.target

        if (target.className === 'go-notify') {
            confirm.classList.add('none')
            backdrop.classList.add('none')
            gonofity.classList.add('none')
            callback();
        }
    })
}

document.addEventListener("DOMContentLoaded", function (event) {
    GetWaitNofityDatas();
    GetCCWaitNofityDatas();
});

function GetWaitNofityDatas() {
    document.getElementsByName("step1-1").forEach((e) => {
        if (!e.classList.contains("none")) {
            e.classList.add("none");
        }
    });

    postData('/CUSCRM/CUSCRMTX001/GetWaitNofity', null, (myJson) => {
        let tbodyWait = document.getElementById('WaitNotify').getElementsByTagName('tbody')[0];
        tbodyWait.innerHTML = "";
        if (myJson != null) {
            document.getElementsByName("step1-1").forEach((e) => { e.classList.remove("none"); });
            document.getElementById('WaitNotifyCount').innerHTML = "已立案未通知清單 共：" + myJson.length + " 筆";
            for (let i = 0; i < myJson.length; i++) {
                var item = myJson[i];
                let newRow = tbodyWait.insertRow();
                let newHidden = document.createElement("input");
                newHidden.setAttribute("type", "hidden");
                newHidden.setAttribute("name", "WaitNotifyCRMENo");
                newHidden.value = item.No;
                let newCell = newRow.insertCell();
                let newText = document.createTextNode((i + 1).toString());
                newCell.appendChild(newText);
                newCell.appendChild(newHidden);

                newCell = newRow.insertCell();
                newText = document.createTextNode(item.Type + '(' + item.TypeName + ')');
                newCell.appendChild(newText);

                newCell = newRow.insertCell();
                newText = document.createTextNode(item.No);
                newCell.appendChild(newText);

                newCell = newRow.insertCell();
                newText = document.createTextNode(item.CreatorName);
                newCell.appendChild(newText);

                newCell = newRow.insertCell();
                newText = document.createTextNode(convertTimestampToDateTime(item.CreateTime));
                newCell.appendChild(newText);

            }
            NotifyRowClick(document.getElementById('WaitNotify'));
        }
    });
}


function GetCCWaitNofityDatas() {
    document.getElementsByName("step1-1").forEach((e) => {
        if (!e.classList.contains("none")) {
            e.classList.add("none");
        }
    });

    postData('/CUSCRM/CUSCRMTX002/GetCCWaitNofity', null, (myJson) => {
        let tbodyWait = document.getElementById('CCWaitNotify').getElementsByTagName('tbody')[0];
        tbodyWait.innerHTML = "";
        if (myJson != null) {
            document.getElementsByName("step1-1").forEach((e) => { e.classList.remove("none"); });
            document.getElementById('CCWaitNotifyCount').innerHTML = "CC已立案未通知申訴人清單 共：" + myJson.length + " 筆";
            for (let i = 0; i < myJson.length; i++) {
                var item = myJson[i];
                let newRow = tbodyWait.insertRow();
                let newHidden = document.createElement("input");
                newHidden.setAttribute("type", "hidden");
                newHidden.setAttribute("name", "WaitNotifyCRMENo");
                newHidden.value = item.No;
                let newCell = newRow.insertCell();
                let newText = document.createTextNode((i + 1).toString());
                newCell.appendChild(newText);
                newCell.appendChild(newHidden);

                newCell = newRow.insertCell();
                newText = document.createTextNode(item.Type + '(' + item.TypeName + ')');
                newCell.appendChild(newText);

                newCell = newRow.insertCell();
                newText = document.createTextNode(item.No);
                newCell.appendChild(newText);

                newCell = newRow.insertCell();
                newText = document.createTextNode(item.CreatorName);
                newCell.appendChild(newText);

                newCell = newRow.insertCell();
                newText = document.createTextNode(convertTimestampToDateTime(item.CreateTime));
                newCell.appendChild(newText);

            }
            CCNotifyRowClick(document.getElementById('CCWaitNotify'));
        }
    });
}


function postData(url, json, callbackFuncs, callbackFunse, callbackFuncf) {
    let formData = new FormData();
    if (json != null) {
        Object.keys(json).forEach((key) => {
            formData.append(key, json[key]);
        });
    }
    if (token != null && token != undefined) {
        formData.append("__RequestVerificationToken", token);
    }

    useBlockUI();
    fetch( url, {
        method: "POST",
        body: formData
    })
    .then(function (response) {
        if (response.url.match(redirectUrl) == redirectUrl) {
            appendMessage('連線逾時，自動轉回登入頁面中，請等待!');
            window.location.href = GetRedirectUrl(response.url, redirectUrl);
        }
        else {
            return response.json();
        }
    })
    .then(function (myJson) {
        callbackFuncs(myJson);
    })
    .catch(error => {
        if (callbackFunse != null) {
            callbackFunse(error);
        }
        else {
            appendMessage(error);
        }
    })
    .finally(() => {
        if (callbackFuncf != null) {
            callbackFuncf();
        }
        $.unblockUI();
    });

}

function PolicyDataToJsonArray() {
    let tempArray = [];
    if (PolicySet.size > 0) {
        let PolicyDatas = Array.from(PolicySet);
        for (let i = 0; i < PolicyDatas.length; i++) {
            tempArray.push(JSON.parse(PolicyDatas[i]));
        }
    }
    return tempArray;
}

function GetCheckedYesValue(elementName) {
    let tempResult = ''
    if (document.getElementById(elementName).checked == true) {
        tempResult = 'Y';
    }
    return tempResult;
}

function PolicyDatasGetComplainants() {
    Complainants.clear();
    let PolicyDatas = PolicyDataToJsonArray();
    for (let i = 0; i < PolicyDatas.length; i++) {
        let item = PolicyDatas[i];
        AddComplainants(item.Owner);
        AddComplainants(item.Insured);
    }
}

function BindSelectPolicy() {
    let pdata = PolicyDataToJsonArray();
    if (pdata.length > 0) {
        document.getElementById('SelectPolicyNoDatas').classList.remove("none");
        document.getElementById('SelectPolicyNoDatasTitle').classList.remove("none");
        document.getElementById('SelectPolicyNoDatasHr').classList.remove("none");
    }
    else {
        document.getElementById('SelectPolicyNoDatas').classList.add("none");
        document.getElementById('SelectPolicyNoDatasTitle').classList.add("none");
        document.getElementById('SelectPolicyNoDatasHr').classList.add("none");
    }
    BindPolicyDatas(document.getElementById('SelectPolicyNoDatas'), pdata);
}

function RemoveExistsPolicyDatas(tableName) {
    let rows = document.getElementById(tableName).rows;
    for (let i = 1; i < rows.length; i++) {
        let chkObj = rows[i].cells[0].childNodes[0];
        if (chkObj.nodeName == 'INPUT') {
            if (chkObj.hasAttribute('data-value')) {
                let dataValue = chkObj.getAttribute('data-value');
                PolicySet.delete(dataValue);
            }
        }
    }
}

function CheckPolicyNoNotClosed() {
    if (PolicyNo.value.trim() != '') {
        postData('/CUSCRM/CUSCRMTX001/CheckNotClosedPolicyNo', { policyNo: PolicyNo.value.trim() }, (myJson) => {
            if (myJson == true) {
                confirm('查該保單尚有未結案之服務案件，請確認是否要繼續立案?', (confirmed) => {
                    if (confirmed) {
                        QueryPolicyNoDatas();
                    }
                }, null, '否,退出');
            }
            else {
                QueryPolicyNoDatas();
            }
        });
    }
    else {
        QueryPolicyNoDatas();
    }
}

function CheckPolicyNoNotClosedYouself() {
    let po_no2 = document.getElementById('PolicyNo2').value.trim();
    if (po_no2 != '') {
        postData('/CUSCRM/CUSCRMTX001/CheckNotClosedPolicyNo', { policyNo: document.getElementById('PolicyNo2').value.trim() }, (myJson) => {
            if (myJson == true) {
                confirm('查該保單尚有未結案之服務案件，請確認是否要繼續立案?', (confirmed) => {
                    if (confirmed) {
                        ShowYouselfContent();
                    }
                }, null, '否,退出');
            }
            else {
                ShowYouselfContent();
            }
        });
    }
    else {
        ShowYouselfContent();
    }
    
}

function QueryPolicyNoDatas() {
    postData('/CUSCRM/CUSCRMTX001/CheckPolicyNo', { policyNo: PolicyNo.value }, (myJson) => {
        if (myJson.result == true) {
            document.getElementsByName("step1").forEach((e) => { e.classList.add("none"); });
            document.getElementsByName("step1-1").forEach((e) => { e.classList.add("none"); });
            document.getElementsByName("step2").forEach((e) => { e.classList.add("none"); });
            document.getElementsByName("step2-3").forEach((e) => { e.classList.add("none"); });
            if (myJson.datas.length == 0) {
                document.getElementsByName("step2-4").forEach((e) => { e.classList.remove("none"); });
            }
            else if (myJson.datas.length == 1) {
                $.unblockUI();
                var data = myJson.datas[0];
                GetPolicyDatasByOwnerID(data.OwnerID);
            }
            else {
                document.getElementsByName("step2-1").forEach((e) => { e.classList.remove("none"); });
                var tbodyRef = document.getElementById('DuplicatePolicyNo').getElementsByTagName('tbody')[0];
                for (let i = 0; i < myJson.datas.length; i++) {
                    var item = myJson.datas[i];
                    let newRow = tbodyRef.insertRow();
                    let newRadio = document.createElement("input");
                    newRadio.setAttribute("type", "radio");
                    newRadio.setAttribute("name", "CheckedOwner");
                    newRadio.className = "form-check-input";
                    newRadio.value = item.OwnerID;
                    let newCell = newRow.insertCell();
                    newCell.appendChild(newRadio);

                    newCell = newRow.insertCell();
                    let newPolicyNoText = document.createTextNode(item.PolicyNo);
                    newCell.appendChild(newPolicyNoText);

                    newCell = newRow.insertCell();
                    let newCompanyText = document.createTextNode(item.CompanyName);
                    newCell.appendChild(newCompanyText);

                    newCell = newRow.insertCell();
                    let newOwnerText = document.createTextNode(item.Owner);
                    newCell.appendChild(newOwnerText);

                    newCell = newRow.insertCell();
                    let newInsuredText = document.createTextNode(item.Insured);
                    newCell.appendChild(newInsuredText);
                }
                RowClick(document.getElementById('DuplicatePolicyNo'));
            }
        }
        else {
            appendMessage('查無接續人，請先派發服務人員');
        }
    });
}

function ShowYouselfContent() {
    let ownerName = document.getElementById('OwnerName').value.trim();
    let insuredName = document.getElementById('InsuredName').value.trim();
    let issDate = document.getElementById('IssDate').value.trim();
    let policyNo2 = document.getElementById('PolicyNo2').value.trim();

    PolicySet.clear();
    PolicySet.add(JSON.stringify({ Owner: ownerName, Insured: insuredName, IssDate: issDate, PolicyNo: policyNo2, VMCode: agentInfo.vm_code, VMName: agentInfo.vm_name, SMCode: agentInfo.sm_code, SMName: agentInfo.sm_name, CenterCode: agentInfo.center_code, CenterName: agentInfo.center_name, WCCode: agentInfo.wc_center, WCName: agentInfo.wc_center_name, AgentCode: agentInfo.agent_code_vlife, AgentName: agentInfo.name, VMLeaderID: agentInfo.vm_leader_id, VMLeader: agentInfo.vm_leader, SMLeaderID: agentInfo.sm_id, SMLeader: agentInfo.sm_leader, CenterLeaderID: agentInfo.administrat_id, CenterLeader: agentInfo.admin_name }));
    document.getElementsByName("step2-4").forEach((e) => { e.classList.add("none"); });
    stepFlag = "step2-4";
    ShowTypeContent();
   
}
