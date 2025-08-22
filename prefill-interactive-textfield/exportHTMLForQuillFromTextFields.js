// Function to retrieve content from a text frame and convert it to HTML with paragraph styles and nested list detection
function exportTextFrameToHTML(textFrame) {
    var htmlContent = "";
    var inList = false;
    var currentListType = null;
    var currentLevel = 0;

    var paragraphStyleMap = {
        "H1": "h1",
        "H2": "h2",
        "H3": "h3",
        "H4": "h4",
        "P": "p",
        "[No Paragraph Style]": "p"
    };

    for (var i = 0; i < textFrame.paragraphs.length; i++) {
        var paragraph = textFrame.paragraphs[i];
        var appliedStyleName = paragraph.appliedParagraphStyle.name;
        var listType = paragraph.bulletsAndNumberingListType;
        var listLevel = listType !== ListType.NO_LIST ? paragraph.numberingLevel : 0;

        if (listType == ListType.BULLET_LIST || listType == ListType.NUMBERED_LIST) {
            if (!inList || listType !== currentListType || listLevel !== currentLevel) {
                if (inList && listLevel < currentLevel) {
                    for (var j = 0; j < currentLevel - listLevel; j++) {
                        htmlContent += "</ol>";
                    }
                }
                if (!inList || listType !== currentListType) {
                    htmlContent += "<ol>";
                }
                inList = true;
            }
            currentListType = listType;

            var listDataAttr = (listType == ListType.NUMBERED_LIST) ? "ordered" : "bullet";
            var indentClass = listLevel > 1 ? ' class="ql-indent-' + listLevel + '"' : "";
            htmlContent += "<li data-list=\"" + listDataAttr + "\"" + indentClass + ">";
        } else {
            if (inList) {
                htmlContent += "</li>";
                for (var j = 0; j < currentLevel; j++) {
                    htmlContent += "</ol>";
                }
                inList = false;
                currentListType = null;
                currentLevel = 0;
            }
            var htmlTag = paragraphStyleMap[appliedStyleName] || "p";
            htmlContent += "<" + htmlTag + ">";
        }

        for (var j = 0; j < paragraph.characters.length; j++) {
            var character = paragraph.characters[j];
            var charContent = character.contents;
            if (character.appliedCharacterStyle.name != "[None]") {
                if (character.appliedCharacterStyle.name == "Bold") {
                    charContent = "<strong>" + charContent + "</strong>";
                }
                if (character.appliedCharacterStyle.name == "Italic") {
                    charContent = "<em>" + charContent + "</em>";
                }
                if (character.appliedCharacterStyle.name == "Underline") {
                    charContent = "<u>" + charContent + "</u>";
                }
            }
            htmlContent += charContent;
        }

        if (inList) {
            htmlContent += "</li>";
        } else {
            htmlContent += "</" + htmlTag + ">";
        }

        currentLevel = listLevel;
    }
    if (inList) {
        for (var j = 0; j < currentLevel; j++) {
            htmlContent += "</ol>";
        }
    }
    return htmlContent;
}


// Function to iterate over all text frames and collect their HTML content if they have a script label
function exportAllTextFieldsToHTML() {
    var allItems = app.activeDocument.allPageItems;
    var result = "";  // String to store the script labels and HTML content
    var first = true;

    result += "{";  // Start JSON-like object

    for (var i = 0; i < allItems.length; i++) {

        if (allItems[i] instanceof TextBox && allItems[i].textFrames.length > 0) {

            var elementName = allItems[i].name || "Unnamed";  // Use 'Unnamed' if no name is set
            var textContent = "";

            var htmlContent = ""

            // Iterate through all textFrames within the TextBox
            for (var j = 0; j < allItems[i].textFrames.length; j++) {

                var _html = exportTextFrameToHTML(allItems[i].textFrames[j]); // Convert to HTML
                htmlContent += _html;
            }

            // Escape special characters in the HTML content
            var escapedHTMLContent = htmlContent
                .replace(/\\/g, '\\\\')   // Escape backslashes
                .replace(/"/g, '\\"')     // Escape double quotes
                .replace(/\n/g, '<br>')    // Escape newlines
                .replace(/\r/g, '')       // Escape carriage returns
                .replace(/\t/g, '');      // Escape tabs

            if (!first) {
                result += ",";  // Add a comma before every item except the first one
            }
            result += '"' + elementName + '": "' + escapedHTMLContent + '"';  // Add escaped content
            first = false;
        }
    }

    result += "}";  // Close JSON-like object

    return result;
}


// Function to save the HTML content as a JSON-like file on the disk
function saveHTMLToJSONFile(jsonContent, exportDetails) {
    var jsonFilePath = exportDetails.folder + "/" + exportDetails.fileName + "_html.json";
    var file = new File(jsonFilePath);

    if (file.open("w", "TEXT", "????")) {
        file.encoding = "UTF-8";
        file.lineFeed = "Unix";
        file.write(jsonContent);  // Manually serialize JSON-like content
        file.close();
        return file.fsName;
    } else {
        alert("Failed to create HTML JSON file");
    }
}
