document.getElementById('convertButton').addEventListener('click', function () {
    const fileInput = document.getElementById('fileInput');
    if (fileInput.files.length === 0) {
        alert("Bitte wähle eine XML-Datei aus!");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function (event) {
        const xmlContent = event.target.result;
        const jsonData = parseXML(xmlContent);
        const excelData = formatJsonToTable(jsonData);
        downloadExcel(excelData);
    };

    reader.readAsText(file);
});

function parseXML(xmlString) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "text/xml");
    const profiles = xmlDoc.getElementsByTagName("TransferProfile");
    const data = [];

    for (let i = 0; i < profiles.length; i++) {
        let profile = profiles[i];
        let row = {
            ProfileGuid: getText(profile, "ProfileGuid"),
            ProfileType: getText(profile, "ProfileType"),
            Name: getText(profile, "Name"),
            DisplayText: getText(profile, "DisplayText"),
            IsActive: getText(profile, "IsActive"),
            RemoteFolderPath: getText(profile, "RemoteFolderPath"),
            LocalFolderPath: getText(profile, "LocalFolderPath"),
            UsePathAuthentication: getText(profile, "UsePathAuthentication"),
            PathUsername: getText(profile, "PathUsername"),
            PathPassword: getText(profile, "PathPassword"),
            SearchPattern: getText(profile, "SearchPattern"),
            Recursive: getText(profile, "Recursive"),
            DeleteFolderIfEmpty: getText(profile, "DeleteFolderIfEmpty"),
            KeepFolders: getText(profile, "KeepFolders")
        };
        data.push(row);
    }
    return data;
}

function getText(xmlNode, tagName) {
    let node = xmlNode.getElementsByTagName(tagName)[0];
    return node ? node.textContent : "";
}

function formatJsonToTable(jsonData) {
    return XLSX.utils.json_to_sheet(jsonData);
}

function downloadExcel(worksheet) {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "TransferProfiles");

    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "TransferProfiles.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
