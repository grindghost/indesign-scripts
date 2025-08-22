// exportHTMLFromTextFrames.jsx

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
        var listLevel = 0;

        if (listType == ListType.BULLET_LIST || listType == ListType.NUMBERED_LIST) {
            listLevel = paragraph.numberingLevel;
        }

        if (listType == ListType.BULLET_LIST || listType == ListType.NUMBERED_LIST) {
            if (!inList) {
                htmlContent += (listType == ListType.BULLET_LIST) ? "<ul>\n" : "<ol>\n";
                inList = true;
            } else if (listType !== currentListType || listLevel > currentLevel) {
                htmlContent += (listType == ListType.BULLET_LIST) ? "<ul>\n" : "<ol>\n";
            } else if (listLevel < currentLevel) {
                for (var j = 0; j < currentLevel - listLevel; j++) {
                    htmlContent += (currentListType == ListType.BULLET_LIST) ? "</ul>\n" : "</ol>\n";
                }
            }

            currentListType = listType;

            var listItemPrefix = "";
            if (listType == ListType.NUMBERED_LIST) {
                listItemPrefix = paragraph.bulletsAndNumberingResultText + " ";
            }

            // htmlContent += "<li style=\"list-style-type: none;\">" + listItemPrefix;
            htmlContent += "<li>" + listItemPrefix;
        } else {
            if (inList) {
                htmlContent += "</li>\n";
                for (var j = 0; j < currentLevel; j++) {
                    htmlContent += (currentListType == ListType.BULLET_LIST) ? "</ul>\n" : "</ol>\n";
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
                    charContent = "<b>" + charContent + "</b>";
                }
                if (character.appliedCharacterStyle.name == "Italic") {
                    charContent = "<i>" + charContent + "</i>";
                }
            }

            htmlContent += charContent;
        }

        if (inList) {
            htmlContent += "</li>\n";
        } else {
            var htmlTag = paragraphStyleMap[appliedStyleName] || "p";
            htmlContent += "</" + htmlTag + ">\n";
        }

        currentLevel = listLevel;
    }

    if (inList) {
        for (var j = 0; j < currentLevel; j++) {
            htmlContent += (currentListType == ListType.BULLET_LIST) ? "</ul>\n" : "</ol>\n";
        }
    }

    return htmlContent;
}

// Function to iterate over all text frames and collect their HTML content if they have a script label
function exportAllTextFramesToHTML() {
    var allItems = app.activeDocument.allPageItems;
    var result = "";  // String to store the script labels and HTML content
    var first = true;

    result += "{\n";  // Start JSON-like object

    for (var i = 0; i < allItems.length; i++) {
        if (allItems[i] instanceof TextFrame && allItems[i].label) {
            var htmlContent = exportTextFrameToHTML(allItems[i]);  // Convert to HTML
            // Escape special characters in the HTML content
            var escapedHTMLContent = htmlContent
                .replace(/\\/g, '\\\\')   // Escape backslashes
                .replace(/"/g, '\\"')     // Escape double quotes
                .replace(/\n/g, '\\n')    // Escape newlines
                .replace(/\r/g, '\\r')    // Escape carriage returns
                .replace(/\t/g, '\\t');   // Escape tabs

            if (!first) {
                result += ",\n";  // Add a comma before every item except the first one
            }
            result += '"' + allItems[i].label + '": "' + escapedHTMLContent + '"';  // Add escaped content
            first = false;
        }
    }

    result += "\n}";  // Close JSON-like object

    return result;
}


// Function to iterate over all text frames and collect their HTML content if they have a script label
function _exportAllTextFramesToHTML() {
    var allItems = app.activeDocument.allPageItems;
    var result = "";  // String to store the script labels and HTML content
    var first = true;

    result += "{\n";  // Start JSON-like object

    for (var i = 0; i < allItems.length; i++) {
        if (allItems[i] instanceof TextFrame && allItems[i].label) {
            var htmlContent = exportTextFrameToHTML(allItems[i]);  // Convert to HTML
            if (!first) {
                result += ",\n";  // Add a comma before every item except the first one
            }
            result += '"' + allItems[i].label + '": ' + '"' + htmlContent.replace(/"/g, '\\"') + '"';  // Escape quotes
            first = false;
        }
    }

    result += "\n}";  // Close JSON-like object

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
