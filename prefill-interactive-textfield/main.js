// Include the external file for HTML export functionality
#include "exportHTMLFromTextFrames.js" // To export associated text frames as HTML
#include "exportHTMLForQuillFromTextFields.js" // To export text fields content as HTML

// Function to integrate HTML export into the existing workflow
function integrateHTMLExport(exportDetails) {
    // Step 1: Extract text frames with script labels and convert them to HTML
    // var htmlJSONContent = exportAllTextFramesToHTML();

    // ...Or extract the content of the text field as html...
    var htmlJSONContent = exportAllTextFieldsToHTML()

    // Step 2: Save the HTML content as a JSON file on the disk
    var htmlJSONFilePath = saveHTMLToJSONFile(htmlJSONContent, exportDetails);

    return htmlJSONFilePath;  // Return the path of the newly created JSON file
}

// Helper function to repeat a string
function repeatString(str, count) {
    var result = "";
    for (var i = 0; i < count; i++) {
        result += str;
    }
    return result;
}

// Helper function to get the number of spaces
function getIndentSpaces(indentInInches, fontSize) {
    var pointsPerInch = 72;
    var avgCharWidth = fontSize / 2; // Approximate width in points
    return Math.round((indentInInches * pointsPerInch) / avgCharWidth);
}

// Function to show a dialog for renaming the PDF file and selecting the export folder
function getExportDetails() {
    // Default file name
    var defaultFileName = app.activeDocument.name.replace(/\.[^\.]+$/, ""); // Remove file extension

    // Create the dialog window
    var dialog = new Window("dialog", "Export Settings");

    // Add a group for the file name entry
    var fileGroup = dialog.add("group");
    fileGroup.add("statictext", undefined, "PDF File Name:");
    var fileNameInput = fileGroup.add("edittext", undefined, defaultFileName);
    fileNameInput.characters = 40; // Set the length of the text entry

    // Add the "OK" and "Cancel" buttons
    var buttonGroup = dialog.add("group");
    var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });
    var cancelButton = buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });

    // Show the dialog and capture the user's action
    if (dialog.show() == 1) {
        var folder = Folder.selectDialog("Select a folder to export PDF and JSON files");
        if (folder) {
            return {
                folder: folder,
                fileName: fileNameInput.text
            };
        } else {
            alert("Export canceled.");
            return null;
        }
    } else {
        alert("Export canceled.");
        return null;
    }
}


// Function to extract the name and content of each text field and export as PDF
function extractTextFieldsAndExportPDF() {
    // Get user input for the file name and export folder
    var exportDetails = getExportDetails();
    if (!exportDetails) return; // If the user cancels, stop execution

    // Set both horizontal and vertical rulers to inches
    app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.INCHES;
    app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.INCHES;


    // Integrate the HTML export and get the path of the HTML JSON file
    var htmlJSONFilePath = integrateHTMLExport(exportDetails);

    var doc = app.activeDocument;
    var interactiveElements = doc.allPageItems;
    var result = {};  // Object to store the name and content
    var tempFrames = []; // Array to store temporary frames for cleanup

    try {
        // Process each text box and save contents
        for (var i = 0; i < interactiveElements.length; i++) {
            var element = interactiveElements[i];

            // Check if the element is a TextBox and contains textFrames
            if (element instanceof TextBox && element.textFrames.length > 0) {
                var elementName = element.name || "Unnamed";  // Use 'Unnamed' if no name is set
                var textContent = "";

                // Iterate through all textFrames within the TextBox
                for (var j = 0; j < element.textFrames.length; j++) {
                    var textFrame = element.textFrames[j];

                    // Create a temporary duplicate of the text frame
                    // Place it outside the visible area
                    var tempFrame = textFrame.duplicate();
                    tempFrames.push(tempFrame); // Store for cleanup later

                    // Move the duplicate frame way off the page
                    tempFrame.move([10000, 10000]);

                    // Now convert bullets to text in the duplicate frame
                    var textRange = tempFrame.texts[0];
                    textRange.convertBulletsAndNumberingToText();

                    // Start with empty content or newline based on position
                    if (j == 0) {
                        var frameContent = "";
                    } else {
                        var frameContent = "\n";
                    }

                    // Process each paragraph to capture indents and converted text
                    for (var p = 0; p < textRange.paragraphs.length; p++) {
                        var paragraph = textRange.paragraphs[p];
                        var indentValue = paragraph.leftIndent || 0; // Get left indent

                        var indentSpaces = getIndentSpaces(indentValue, 9); // Helvetica 12 pt
                        var indentString = repeatString(" ", indentSpaces);

                        frameContent += indentString + paragraph.contents;
                    }

                    // Ensure frameContent is a string before appending it
                    if (typeof frameContent === "string") {
                        textContent += frameContent;  // Append content from each frame
                    }
                }

                // Add the element name and its text content to the result object
                result[elementName] = textContent;
            }
        }

        // Clean up all the temporary frames
        for (var k = 0; k < tempFrames.length; k++) {
            tempFrames[k].remove();
        }

    } catch (e) {
        // Clean up even if there's an error
        for (var k = 0; k < tempFrames.length; k++) {
            try { tempFrames[k].remove(); } catch(err) {}
        }
        alert("Error processing document: " + e);
    }

    // Serialize the result to JSON format using your custom serializer
    var jsonString = jsonSerialize(result);

    // Save the JSON to a file on the disk
    var jsonFilePath = saveToJSONFile(jsonString, exportDetails);

    // Export the document as an interactive PDF
    exportToPDF(function (pdfPath) {
        // When the export is complete, trigger the Python script with the PDF path and JSON file path
        triggerPythonScript(pdfPath, jsonFilePath, htmlJSONFilePath, exportDetails);
    }, exportDetails);
}

// Function to serialize object to JSON format with handling for line breaks, carriage returns, and tab characters
function jsonSerialize(obj) {
    var json = '{\n';
    var first = true;

    for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
            if (!first) {
                json += ',\n';  // Add a comma before every key except the first one
            }

            // Replace carriage returns (\r) with two line breaks (\n\n), line breaks (\n) with one line break (\n),
            // and tabs (\t) with a single space.
            var formattedText = obj[key]
                .replace(/\r/g, '\\n')  // Replace carriage returns with \n\n
                .replace(/\n/g, '\\n')     // Replace line breaks with \n
                .replace(/\t/g, ' ');      // Replace tabs with a single space

            json += '  "' + key + '": ' + '"' + formattedText.replace(/"/g, '\\"') + '"';  // Escape quotes
            first = false;
        }
    }

    json += '\n}';
    return json;
}

// Function to save the extracted data as a JSON file on the desktop
function saveToJSONFile(jsonString, exportDetails) {
    var jsonFilePath = exportDetails.folder + "/" + exportDetails.fileName + ".json";
    var file = new File(jsonFilePath);

    if (file.open("w", "TEXT", "????")) {  // Open file for writing with specific type
        file.encoding = "UTF-8";  // Ensure UTF-8 encoding for special characters
        file.lineFeed = "Unix";  // Use Unix style line feeds
        file.write(jsonString);  // Save the manually serialized JSON
        file.close();
        return file.fsName;
        // alert("JSON file saved to desktop: " + desktopPath);
    } else {
        // alert("Failed to create JSON file");
    }
}


// Function to export the document as an interactive PDF to the desktop
function exportToPDF(callback, exportDetails) {
    var pdfFile = new File(exportDetails.folder + "/" + exportDetails.fileName + ".pdf");

    // Disable automatic viewing of the PDF after export
    app.interactivePDFExportPreferences.viewPDF = false;

    // View options
    app.interactivePDFExportPreferences.pdfDisplayTitle = PdfDisplayTitleOptions.DISPLAY_DOCUMENT_TITLE;
    app.interactivePDFExportPreferences.pdfMagnification = PdfMagnificationOptions.FIT_PAGE;
    app.interactivePDFExportPreferences.pdfPageLayout = PageLayoutOptions.SINGLE_PAGE;

    // Create tagged PDF (if I need it)
    // app.interactivePDFExportPreferences.includeStructure = true;
    // app.interactivePDFExportPreferences.usePDFStructureForTabOrder = true;

    // PDF Export Preset for interactive PDFs
    var preset = app.pdfExportPresets.item("[Smallest File Size]"); // This is for a print-based export
    var exportFormat = ExportFormat.INTERACTIVE_PDF; // Ensure the correct export format

    app.activeDocument.exportFile(exportFormat, pdfFile, false, preset);
    // Check every 500ms whether the file exists
    waitForFile(pdfFile, function() {
        // Execute the callback once export is done
        callback(pdfFile.fsName);  // Full path of the PDF
    });
}

// Function to check if the file exists
function waitForFile(file, callback) {
    var maxRetries = 100;  // Retry up to 100 times (adjust if necessary)
    var retries = 0;

    // Function to check file existence
    function checkFile() {
        if (file.exists) {
            callback();
        } else {
            retries++;
            if (retries < maxRetries) {
                $.sleep(500);  // Wait 500ms before checking again
                checkFile();   // Retry
            } else {
                alert("Failed to detect PDF file after several attempts.");
            }
        }
    }

    // Start checking
    checkFile();
}


// Function to trigger the Python script
function triggerPythonScript(pdfPath, jsonFilePath, htmlJSONFilePath, exportDetails) {
    // Get the path of the current ExtendScript file
    var scriptFolder = File($.fileName).path;  // Get the folder of the running ExtendScript

    // Path to the Python script in the same folder as the ExtendScript
    var pythonScriptPath = File(scriptFolder + "/prefill_pdf.py").fsName;  // Get the absolute path

    // Use the system Python path or specify your own Python executable if needed
    var pythonPath = "/Library/Frameworks/Python.framework/Versions/3.10/bin/python3";  // Update this to your Python path

    // Use the same export folder for the completion file
    var completionFile = new File(exportDetails.folder + "/python_complete.txt");
    var completionFilePath = completionFile.fsName;  // Get the full absolute path

    // Construct the AppleScript to call Python with the necessary arguments
    var appleScript = 'do shell script "' + pythonPath + ' \\"' + pythonScriptPath + '\\" \\"' + pdfPath + '\\" \\"' + jsonFilePath + '\\" \\"' + htmlJSONFilePath + '\\" \\"' + completionFilePath + '\\""';

    app.doScript(appleScript, ScriptLanguage.APPLESCRIPT_LANGUAGE);

    // Check for the completion file created by Python
    waitForCompletionFile(completionFilePath, pdfPath, jsonFilePath, htmlJSONFilePath);
}

// Function to wait for the Python script to complete by checking for the completion file
function waitForCompletionFile(completionFilePath, pdfPath, jsonFilePath, htmlJSONFilePath) {
    var completionFile = new File(completionFilePath);
    var pdfFile = new File(pdfPath);
    var jsonFile = new File(jsonFilePath);
    var htmlJsonFile = new File(htmlJSONFilePath);

    var maxRetries = 100;  // Retry up to 100 times (adjust if necessary)
    var retries = 0;

    function checkFile() {
        if (completionFile.exists) {
            alert("Python script completed successfully!");
            completionFile.remove();  // Clean up the completion file after confirming completion
            pdfFile.remove();
            jsonFile.remove();
            htmlJsonFile.remove();

        } else {
            retries++;
            if (retries < maxRetries) {
                $.sleep(500);  // Wait 500ms before checking again
                checkFile();
            } else {
                alert("Python script did not complete within the expected time.");
            }
        }
    }

    checkFile();
}

// Run the function
extractTextFieldsAndExportPDF();
