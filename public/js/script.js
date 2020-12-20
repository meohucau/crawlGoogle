
function validateFrom() {
    var companyList = document.getElementById("companyNameList").value.split("\n");;
    var drugList = document.getElementById("drugNameList").value.split("\n");;
    var numberResult = document.getElementById("numberResult").value;
    if (companyList == "") {
        alert("Please input company list");
        return false;
    }
    if (drugList == "") {
        alert("Please input drug list");
        return false;
    }
    if (numberResult == null || numberResult == "") {
        alert("Please input number result");
        return false;
    }
    if(companyList.length !== drugList.length) {
        alert("Drug list number is not equals with company list")
        return false;
    }
}
