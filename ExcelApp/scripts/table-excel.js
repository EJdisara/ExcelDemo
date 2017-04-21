
function edit(no) {
    document.getElementById("td_" + no).style.display = "none";
    document.getElementById("editTd_" + no).style.display = "none";
    document.getElementById("div_" + no).style.display = "block";

}

function saveColumn(no)
{
    var inputVal = document.getElementById("data_"+no).value;

    document.getElementById("td_" + no).innerHTML = inputVal;

    document.getElementById("td_" + no).style.display = "block";
    document.getElementById("editTd_" + no).style.display = "block";
    document.getElementById("div_" + no).style.display = "none";
}

function cancelColumn(no){
    document.getElementById("td_" + no).style.display = "block";
    document.getElementById("editTd_" + no).style.display = "block";
    document.getElementById("div_" + no).style.display = "none";
}

function addRow(no) {
    document.getElementById("addRow_" + no).style.display = "none";
    document.getElementById("deleteRow_" + no).style.display = "none";
    var tdClass = document.getElementsByClassName("tdColor_" + no), i, len;
    for (i = 0, len = tdClass.length; i < len; i++) {
        tdClass[i].style.background = 'white';
    }

    var btnEditClass = document.getElementsByClassName("btnHide_" + no), i, len;
    for (i = 0, len = btnEditClass.length; i < len; i++) {
        btnEditClass[i].style.display = 'none';
    }
}

function deleteRow(no){
    document.getElementById("row_"+no).outerHTML="";
}


