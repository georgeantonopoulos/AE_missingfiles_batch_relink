(function() { // Wrap in a function to avoid polluting global scope and for cleaner error handling
    try {
        var scriptFile = new File($.fileName);

        // Check if scriptFile and its parent are valid
        if (!scriptFile.exists) {
            alert("Error: Script file ('" + $.fileName + "') does not seem to exist. Please ensure the script is saved and run from After Effects.");
            return;
        }
        var scriptParentFolder = scriptFile.parent;
        if (scriptParentFolder === null) {
            alert("Error: Cannot determine the script's parent folder. The script might be located at the root of a drive. Please place it in a subfolder.");
            return;
        }
        var scriptFolder = scriptParentFolder.fsName;
        var projectItemCount = app.project.numItems;

        alert("Script 'exportMissingFiles.jsx' starting.\n" +
              "Script folder: " + scriptFolder + "\n" +
              "Project items to check: " + projectItemCount);

        var missing = [];
        if (projectItemCount === 0) {
            // This case is handled later by the "No missing footage detected" if missing.length remains 0,
            // or an explicit message if preferred.
            // For now, let the main logic proceed; if missing stays empty, it'll be handled.
        }

        for (var i = 1; i <= projectItemCount; i++) {
            var itm = app.project.item(i);
            // Ensure itm.mainSource exists before checking missingFootagePath
            if (itm instanceof FootageItem && itm.mainSource && itm.mainSource.missingFootagePath) {
                missing.push(itm.name + " ? " + itm.mainSource.missingFootagePath);
            }
        }

        if (missing.length > 0) {
            var outFilePath = scriptFolder + "/missingFiles.txt";
            var outFile = new File(outFilePath);
            outFile.encoding = "UTF-8";

            try {
                if (!outFile.open("w")) {
                    throw new Error("Failed to open file for writing: '" + outFile.name + "'.\nPossible permission issue in folder: " + scriptFolder + ".\nDetails: " + outFile.error);
                }
                for (var j = 0; j < missing.length; j++) {
                    if (!outFile.writeln(missing[j])) {
                        throw new Error("Failed to write line: '" + missing[j] + "' to file '" + outFile.name + "'.\nDetails: " + outFile.error);
                    }
                }
                var closedSuccessfully = outFile.close();
                if (!closedSuccessfully) {
                    // This is a warning; data might have been flushed but the close operation reported an issue.
                    alert("Warning: File '" + outFile.name + "' may not have closed properly, but data was likely written.\nDetails: " + outFile.error);
                }
                // If no errors were thrown during open/write, consider it a success.
                alert("Successfully wrote " + missing.length + " missing footage entries to:\n" + outFile.fsName);

            } catch (fileError) {
                alert("Error during file operation:\n" +
                      "File: " + (outFile.fsName || outFilePath) + "\n" + // Use fsName if available
                      "Error: " + (fileError.message || fileError.toString()));
                // Attempt to close the file if it might be open, ignore errors on this close attempt
                try { if (outFile && typeof outFile.close === 'function') outFile.close(); } catch (e) { /* ignore */ }
            }
        } else {
            if (projectItemCount > 0) {
                alert("No missing footage detected after checking " + projectItemCount + " items.");
            } else {
                alert("The project is empty. No items to check for missing footage.");
            }
        }

    } catch (e) {
        var errorMsg = "An unexpected error occurred in 'exportMissingFiles.jsx':\n";
        errorMsg += "Message: " + (e.message || e.toString()) + "\n";
        if (e.line !== undefined) errorMsg += "Line: " + e.line + "\n"; // e.line might be undefined
        if (e.fileName) errorMsg += "Script File: " + File(e.fileName).displayName + "\n";
        alert(errorMsg);
    }
})(); // Execute the main function 