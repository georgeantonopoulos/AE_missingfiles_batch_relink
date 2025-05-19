//@target aftereffects

(function BatchCollectAndRelink() {
    var mainUndoGroup = "Batch Collect and Relink Footage";
    var collectedRootFolderDest = null;
    var collectionLogFilePath = null;
    var finalLogFilePath = null;
    var searchBaseFolder = null;

    // --- Helper function for recursive search - FINDS ALL MATCHES ---
    function findFilesRecursive(baseFolder, fileNameToFind, foundFilesArray) {
        var items = baseFolder.getFiles();
        if (items) { // Check if getFiles() returned null (e.g. permission issue)
            for (var i = 0; i < items.length; i++) {
                var currentItem = items[i];
                if (currentItem instanceof File && currentItem.name.toLowerCase() === fileNameToFind.toLowerCase()) {
                    if(currentItem.exists) foundFilesArray.push(currentItem); // Add to list if it actually exists
                }
                if (currentItem instanceof Folder) {
                    findFilesRecursive(currentItem, fileNameToFind, foundFilesArray); // Recurse
                }
            }
        }
    }

    try {
        app.beginUndoGroup(mainUndoGroup);

        // --- PART 1: GATHER INFO & COLLECT FILES BY SEARCHING ---
        alert("Part 1: Collect Missing Files by Searching\n\nStep 1: Select your 'missingFiles.txt' file.");
        var missingFilesTxt = File.openDialog("Select your missingFiles.txt", "*.txt");
        if (!missingFilesTxt || !missingFilesTxt.exists) {
            alert("MissingFiles.txt not selected or does not exist. Script will exit.");
            try { app.endUndoGroup(); } catch(e) {}
            return;
        }

        alert("Step 2: Select the TOP-LEVEL FOLDER to SEARCH WITHIN for your missing source files.");
        searchBaseFolder = Folder.selectDialog("Select Base Folder for Recursive Search");
        if (!searchBaseFolder || !searchBaseFolder.exists) {
            alert("Search folder not selected or does not exist. Script will exit.");
            try { app.endUndoGroup(); } catch(e) {}
            return;
        }

        alert("Step 3: Select or Create a DESTINATION folder for the successfully found and COPIED footage.");
        collectedRootFolderDest = Folder.selectDialog("Select/Create Destination Folder for Collected Footage");
        if (!collectedRootFolderDest) {
            alert("No destination folder selected. Script will exit.");
            try { app.endUndoGroup(); } catch(e) {}
            return;
        }
        if (!collectedRootFolderDest.exists) {
            if (confirm("Destination folder: \n" + collectedRootFolderDest.fsName + "\ndoes not exist. Create it?")) {
                if (!collectedRootFolderDest.create()) {
                    alert("Could not create destination folder. Script will exit.");
                    try { app.endUndoGroup(); } catch(e) {}
                    return;
                }
            } else {
                alert("Destination folder not created. Script will exit.");
                try { app.endUndoGroup(); } catch(e) {}
                return;
            }
        }
        
        var collectionLog = "--- File Collection Log (Search Based with Duplicate Handling) ---\nGenerated at: " + new Date().toString() + "\n";
        collectionLog += "Missing Files List: " + missingFilesTxt.fsName + "\n";
        collectionLog += "Search Base Folder: " + searchBaseFolder.fsName + "\n";
        collectionLog += "Collection Destination: " + collectedRootFolderDest.fsName + "\n\n";

        var copiedCount = 0;
        var collectionErrors = 0;
        var skippedMalformedCount = 0;
        var searchNotFoundCount = 0;

        missingFilesTxt.open("r");
        var lineCount = 0;
        var filesToSearch = [];
        while(!missingFilesTxt.eof){
            lineCount++;
            var currentLine = missingFilesTxt.readln();
            if (currentLine.replace(/^\s+|\s+$/g, '') === "") continue;
            var parts = currentLine.split(' ? ');
            if (parts.length < 2) {
                collectionLog += "Skipping Malformed Line (" + lineCount + "): " + currentLine + "\n";
                skippedMalformedCount++;
                continue;
            }
            var itemName = parts[0].replace(/^\s+|\s+$/g, '');
            var originalPathStr = parts.slice(1).join(' ? ').replace(/^\s+|\s+$/g, '');
            var tempFile = new File(originalPathStr); 
            var originalFileName = tempFile.name;
            var originalParentDir = tempFile.parent;
            var originalParentDirName = originalParentDir ? originalParentDir.name : "_root_";
            filesToSearch.push({name: itemName, fileName: originalFileName, parentDirName: originalParentDirName.toLowerCase(), originalPath: originalPathStr}); // Store parentDirName as lowercase for comparison
        }
        missingFilesTxt.close();

        collectionLog += "Found " + filesToSearch.length + " files to process from missingFiles.txt.\n\n";
        alert("Starting search for " + filesToSearch.length + " distinct file entries within: \n" + searchBaseFolder.fsName + "\nThis may take some time...");

        for(var f=0; f < filesToSearch.length; f++){
            var fileEntry = filesToSearch[f];
            collectionLog += "Processing AE Item: '" + fileEntry.name + "', Original File: '" + fileEntry.fileName + "', Original Parent Dir Name: '" + fileEntry.parentDirName + "'\n";
            
            var allFoundSourceFiles = [];
            findFilesRecursive(searchBaseFolder, fileEntry.fileName, allFoundSourceFiles);
            var chosenSourceFile = null;

            if (allFoundSourceFiles.length === 0) {
                collectionLog += "  -> File NOT FOUND by search in '" + searchBaseFolder.displayName + "' (Original full path was: '" + fileEntry.originalPath + "')\n";
                searchNotFoundCount++;
            } else if (allFoundSourceFiles.length === 1) {
                chosenSourceFile = allFoundSourceFiles[0];
                collectionLog += "  -> FOUND unique match at: '" + chosenSourceFile.fsName + "'\n";
            } else { // Multiple matches found
                collectionLog += "  -> WARNING: Multiple (" + allFoundSourceFiles.length + ") potential source files found for '" + fileEntry.fileName + "':\n";
                for (var k=0; k < allFoundSourceFiles.length; k++) {
                    collectionLog += "    - Candidate " + (k+1) + ": " + allFoundSourceFiles[k].fsName + " (Parent: '" + (allFoundSourceFiles[k].parent ? allFoundSourceFiles[k].parent.name : "N/A") + "')\n";
                }

                var parentMatches = [];
                for (var k=0; k < allFoundSourceFiles.length; k++) {
                    var candidateParent = allFoundSourceFiles[k].parent;
                    if (candidateParent && candidateParent.name.toLowerCase() === fileEntry.parentDirName) { // Compare with stored lowercase parentDirName
                        parentMatches.push(allFoundSourceFiles[k]);
                    }
                }

                if (parentMatches.length === 1) {
                    chosenSourceFile = parentMatches[0];
                    collectionLog += "    -> SELECTED candidate based on uniquely matching parent folder name ('" + fileEntry.parentDirName + "'): " + chosenSourceFile.fsName + "\n";
                } else if (parentMatches.length > 1) {
                    chosenSourceFile = parentMatches[0]; 
                    collectionLog += "    -> Multiple candidates matched parent folder name ('" + fileEntry.parentDirName + "'). SELECTED FIRST of these: " + chosenSourceFile.fsName + "\n";
                    collectionLog += "       Review other parent-matching candidates if this is incorrect.\n";
                } else { // No parent folder matches, take the overall first found
                    chosenSourceFile = allFoundSourceFiles[0];
                    collectionLog += "    -> No candidates matched parent folder name ('" + fileEntry.parentDirName + "'). SELECTED FIRST overall candidate found: " + chosenSourceFile.fsName + "\n";
                    collectionLog += "       Review other candidates if this selection is incorrect.\n";
                }
            }

            if (chosenSourceFile) {
                // Proceed with copying chosenSourceFile
                var destinationSubfolder = new Folder(collectedRootFolderDest.fsName + "/" + (new File(fileEntry.originalPath).parent ? new File(fileEntry.originalPath).parent.name : "_root_")); // Use original parent name for subfolder
                if (!destinationSubfolder.exists) {
                    if (!destinationSubfolder.create()) {
                        collectionLog += "    ERROR: Could not create destination subfolder: " + destinationSubfolder.fsName + " for '" + chosenSourceFile.name + "'\n";
                        collectionErrors++;
                        continue; // Skip to next fileEntry in outer loop
                    }
                }

                var destinationFilePath = destinationSubfolder.fsName + "/" + chosenSourceFile.name; // Use name of the CHOSEN source file
                var finalCopiedFile = new File(destinationFilePath);
                var counter = 1;
                var baseName = chosenSourceFile.name.substring(0, chosenSourceFile.name.lastIndexOf('.'));
                if (chosenSourceFile.name.lastIndexOf('.') === -1) baseName = chosenSourceFile.name;
                var extension = chosenSourceFile.name.substring(chosenSourceFile.name.lastIndexOf('.'));
                if (chosenSourceFile.name.lastIndexOf('.') === -1) extension = "";

                while (finalCopiedFile.exists) {
                    var newFileName = baseName + "_" + counter + extension;
                    destinationFilePath = destinationSubfolder.fsName + "/" + newFileName;
                    finalCopiedFile = new File(destinationFilePath);
                    counter++;
                }
                if (counter > 1) {
                    collectionLog += "    Info: Destination file '" + chosenSourceFile.name + "' exists in target subfolder. Renaming copied file to '" + finalCopiedFile.name + "'.\n";
                }

                try {
                    if (chosenSourceFile.copy(destinationFilePath)) {
                        collectionLog += "    COPIED '" + chosenSourceFile.fsName + "' TO: '" + destinationFilePath + "'\n";
                        copiedCount++;
                    } else {
                        collectionLog += "    ERROR: FAILED to copy from '" + chosenSourceFile.fsName + "' TO '" + destinationFilePath + "'. OS Error: '" + chosenSourceFile.error + "'\n";
                        collectionErrors++;
                    }
                } catch (copyError) {
                    collectionLog += "    EXCEPTION during copy of '" + chosenSourceFile.fsName + "': " + copyError.toString() + "\n";
                    collectionErrors++;
                }
            }
            collectionLog += "---\n";
        }

        collectionLog += "\n--- Collection Summary (Search Based with Duplicate Handling) ---\n";
        collectionLog += "Files successfully copied: " + copiedCount + "\n";
        collectionLog += "Files NOT FOUND by search: " + searchNotFoundCount + "\n";
        collectionLog += "Malformed lines in txt: " + skippedMalformedCount + "\n";
        collectionLog += "File copy/IO errors during collection: " + collectionErrors + "\n";
        collectionLog += "Collection phase finished at: " + new Date().toString() + "\n";

        collectionLogFilePath = collectedRootFolderDest.fsName + "/_CollectionLog_SearchDupHandling.txt";
        var collectionLogFile = new File(collectionLogFilePath);
        try {
            collectionLogFile.open("w"); collectionLogFile.encoding = "UTF-8";
            collectionLogFile.writeln(collectionLog);
            collectionLogFile.close();
            alert("File Collection (Search with Duplicate Handling) phase complete. Log saved to:\n" + collectionLogFile.fsName);
        } catch (logWriteError) {
            alert("CRITICAL ERROR writing collection log: " + logWriteError.toString() + "\nPartial log content (first 500 chars):\n"+collectionLog.substring(0,500));
        }
        
        if (copiedCount === 0) {
            var exitMsg = "No files were copied after searching.";
            if (searchNotFoundCount > 0) exitMsg += " Many files were not found or no suitable candidate could be chosen.";
            exitMsg += "\n\nPlease review '" + (collectionLogFile.exists ? collectionLogFile.name : "_CollectionLog_SearchDupHandling.txt") + "' in the destination folder.\n\nRelinking will be skipped.";
            alert(exitMsg);
            app.endUndoGroup(); return;
        }

        // --- PART 2: RELINKING FOOTAGE ---
        alert("Part 2: Relink Footage\n\nAttempting to relink using files in:\n" + collectedRootFolderDest.fsName);

        var relinkedCount = 0;
        var relinkNotFoundCount = 0;
        var alreadyOkCount = 0;
        var relinkErrorsCount = 0;
        var itemsProcessed = 0;
        var relinkLog = "\n\n--- Relinking Log (Post Search-Collection with DupHandling) ---\nGenerated at: " + new Date().toString() + "\n";
        relinkLog += "Using collected footage from: " + collectedRootFolderDest.fsName + "\n\n";

        for (var i = 1; i <= app.project.numItems; i++) {
            var item = app.project.item(i);
            itemsProcessed++;

            if (item instanceof FootageItem && item.mainSource && item.mainSource.missingFootagePath) {
                var originalMissingPathStr = item.mainSource.missingFootagePath;
                var tempOriginalFile = new File(originalMissingPathStr);
                var originalFileName = tempOriginalFile.name;
                var originalParentDirNameForRelink = tempOriginalFile.parent ? tempOriginalFile.parent.name : "_root_";

                var targetSubfolder = new Folder(collectedRootFolderDest.fsName + "/" + originalParentDirNameForRelink);
                var foundNewFile = null;

                if (targetSubfolder.exists) {
                    var testFile = new File(targetSubfolder.fsName + "/" + originalFileName);
                    if (testFile.exists) {
                        foundNewFile = testFile;
                    } else {
                        var baseNameRelink = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
                        if (originalFileName.lastIndexOf('.') === -1) baseNameRelink = originalFileName;
                        var extensionRelink = originalFileName.substring(originalFileName.lastIndexOf('.'));
                        if (originalFileName.lastIndexOf('.') === -1) extensionRelink = "";
                        for (var j = 1; j <= 100; j++) { 
                            var renamedFileName = baseNameRelink + "_" + j + extensionRelink;
                            var testRenamedFile = new File(targetSubfolder.fsName + "/" + renamedFileName);
                            if (testRenamedFile.exists) {
                                foundNewFile = testRenamedFile;
                                relinkLog += "Info: For AE item '"+item.name+"', original file '"+originalFileName+"' was collected as '" + testRenamedFile.name + "' for relinking.\n";
                                break;
                            }
                        }
                    }
                }

                if (foundNewFile && foundNewFile.exists) {
                    try {
                        item.replace(foundNewFile);
                        if (item.mainSource.missingFootagePath) {
                            relinkLog += "ERROR: Relink reported success for '" + item.name + "', but still missing. (Tried: "+ foundNewFile.fsName +")\n";
                            relinkErrorsCount++;
                        } else {
                            relinkLog += "Relinked '" + item.name + "' to '" + foundNewFile.fsName + "'\n";
                            relinkedCount++;
                        }
                    } catch (replaceError) {
                        relinkLog += "EXCEPTION during relink of '" + item.name + "' to '" + foundNewFile.fsName + "': " + replaceError.toString() + "\n";
                        relinkErrorsCount++;
                    }
                } else {
                    relinkLog += "Not Found in Collected: Could not find '" + originalFileName + "' (or variants) for item '" + item.name + "' in collected subfolder '" + (targetSubfolder.exists ? targetSubfolder.displayName : originalParentDirNameForRelink + " [Not Created/Found]") + "'. Original path: " + originalMissingPathStr + "\n";
                    relinkNotFoundCount++;
                }
            } else if (item instanceof FootageItem && item.file && item.file.exists) {
                alreadyOkCount++;
            }
        }

        relinkLog += "\n--- Relinking Summary ---\n";
        relinkLog += "Total Items Checked: " + itemsProcessed + "\n";
        relinkLog += "Successfully Relinked: " + relinkedCount + "\n";
        relinkLog += "Not Found in Collected Folder: " + relinkNotFoundCount + "\n";
        relinkLog += "Already Online: " + alreadyOkCount + "\n";
        relinkLog += "Relink Errors: " + relinkErrorsCount + "\n";
        relinkLog += "Relinking finished: " + new Date().toString() + "\n";

        finalLogFilePath = collectedRootFolderDest.fsName + "/_BatchCollectAndRelink_FullLog_SearchDupHandling.txt";
        var finalLogFile = new File(finalLogFilePath);
        try {
            finalLogFile.open("w"); finalLogFile.encoding = "UTF-8";
            finalLogFile.writeln(collectionLog + relinkLog);
            finalLogFile.close();
            alert("Process Complete. Full log: " + finalLogFile.fsName + 
                  "\n\nRelink Summary:\nRelinked: " + relinkedCount + "; Not Found: " + relinkNotFoundCount + "; Errors: " + relinkErrorsCount);
        } catch (finalLogWriteError) {
            alert("CRITICAL ERROR writing final log: " + finalLogWriteError.toString() + ".\nRelink Summary: Relinked: " + relinkedCount + "; Not Found: " + relinkNotFoundCount + "; Errors: " + relinkErrorsCount);
        }
        app.endUndoGroup();

    } catch (e) {
        var errorMsg = "A CRITICAL SCRIPT ERROR occurred:\n" + 
                       "Message: " + (e.message || e.toString()) + "\n" +
                       "Line: " + (e.line || "N/A") + "\n" +
                       "File: " + (e.fileName ? File(e.fileName).displayName : "BatchCollectAndRelink.jsx") + "\n" +
                       "Stack: " + (e.stack || "N/A") + "\n\n" +
                       "Check for partial logs in: " + (collectedRootFolderDest ? collectedRootFolderDest.fsName : "(destination folder not set)") + "\n";
        if(collectionLogFilePath) errorMsg += "Collection Log (if created): " + collectionLogFilePath + "\n";
        if(finalLogFilePath) errorMsg += "Final Log (if created): " + finalLogFilePath + "\n";
        alert(errorMsg);
        try { if (missingFilesTxt && !missingFilesTxt.closed) missingFilesTxt.close(); } catch (cerr) {}
        try { app.endUndoGroup(); } catch (e2) {}
    }
})(); 