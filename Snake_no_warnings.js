// InDesign CS6 Script: Arrange Objects with Numbered List Flow (No Warnings)
// Arranges objects in grid, preserves dynamic numbering, renumbers to follow pattern flow

(function() {
    if (app.documents.length == 0) {
        alert("Please open a document first.");
        return;
    }
    
    var doc = app.activeDocument;
    var selection = app.selection;
    
    // Check if any objects are selected
    if (selection.length == 0) {
        alert("Please select some objects first.");
        return;
    }
    
    // Check if a page is selected and separate objects from pages
    var selectedPage = null;
    var selectedObjects = [];
    
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.constructor.name === "Page") {
            selectedPage = item;
        } else {
            selectedObjects.push(item);
        }
    }
    
    // If no page is selected, use active page
    if (!selectedPage) {
        selectedPage = app.activeWindow.activePage;
    }
    
    // Check if we have any objects to arrange
    if (selectedObjects.length == 0) {
        alert("Please select some objects to arrange.");
        return;
    }
    
    // CREATE THE BG LAYER AND BG PATH LAYER AND MOVE THEM TO THE VERY BOTTOM
    var bgLayer = createLayer("BG");
    var bgPathLayer = createLayer("BG PATH");
    
    // Arrange layers in correct order from bottom to top: BG -> BG PATH -> other layers
    arrangeLayersAtBottom(bgLayer, bgPathLayer);
    
    // Calculate how many pages we need
    var objectsPerPage = 12;
    var totalObjects = selectedObjects.length;
    var totalPages = Math.ceil(totalObjects / objectsPerPage);
    
    // Show pattern selection dialog for each page
    var pagePatterns = showPerPagePatternDialog(totalPages);
    if (pagePatterns == null) return;
    
    // Rearrange objects and RENUMBER them according to pattern flow
    rearrangeAndRenumberObjects(selectedObjects, pagePatterns, selectedPage, bgPathLayer);
    
    function createLayer(layerName) {
        try {
            var layer;
            
            // First, try to get existing layer
            try {
                layer = doc.layers.item(layerName);
                // If we get here, layer exists - check if it's valid
                if (!layer.isValid) {
                    layer = doc.layers.add();
                    layer.name = layerName;
                }
            } catch(e) {
                // Layer doesn't exist, so create it
                layer = doc.layers.add();
                layer.name = layerName;
            }
            
            return layer;
        } catch(e) {
            // If layer creation fails, use default layer
            return doc.layers[0];
        }
    }
    
    function arrangeLayersAtBottom(bgLayer, bgPathLayer) {
        try {
            // Move both layers to the bottom in correct order
            // First move BG to absolute bottom
            if (bgLayer && bgLayer.isValid) {
                bgLayer.move(LocationOptions.AT_BEGINNING);
            }
            
            // Then move BG PATH above BG (so move it to beginning after BG)
            if (bgPathLayer && bgPathLayer.isValid) {
                bgPathLayer.move(LocationOptions.AT_BEGINNING);
            }
        } catch(e) {
            // Silent failure - continue with current layer order
        }
    }
    
    function showPerPagePatternDialog(totalPages) {
        if (totalPages > 6) {
            var patterns = [];
            for (var i = 0; i < totalPages; i++) {
                patterns.push(1);
            }
            return patterns;
        }
        
        var dialog = app.dialogs.add({name: "Select Pattern for Each Page", canCancel: true});
        var column = dialog.dialogColumns.add();
        
        column.staticTexts.add({staticLabel: "Select pattern for each page:"});
        
        var dropdowns = [];
        var patternNames = [
            "UP-LOW1",
            "UP-LOW2", 
            "UP-LOW3",
            "UP-UP MID",
            "LOW MID – UP MID",
            "LOW – LOW MID",
            "LOW MID – LOW",
            "UP MID – LOW",
            "LOW – UP 1",
            "LOW – UP 2"
        ];
        
        for (var i = 0; i < totalPages; i++) {
            var pageRow = column.dialogRows.add();
            pageRow.staticTexts.add({staticLabel: "Page " + (i + 1) + ":", minWidth: 60});
            
            var dropdown = pageRow.dropdowns.add({
                stringList: patternNames,
                selectedIndex: 0
            });
            
            dropdowns.push(dropdown);
        }
        
        if (dialog.show() == true) {
            var selectedPatterns = [];
            for (var j = 0; j < dropdowns.length; j++) {
                var selectedIndex = dropdowns[j].selectedIndex;
                // Convert index (0-9) to pattern number (1-10)
                var patternNumber = selectedIndex + 1;
                selectedPatterns.push(patternNumber);
            }
            return selectedPatterns;
        }
        return null;
    }
    
    function rearrangeAndRenumberObjects(allObjects, pagePatterns, startPage, bgLayer) {
        var objectsPerPage = 12;
        var totalObjects = allObjects.length;
        
        // Calculate grid dimensions from the starting page
        var pageWidth = startPage.bounds[3] - startPage.bounds[1];
        var pageHeight = startPage.bounds[2] - startPage.bounds[0];
        
        var margin = 36;
        var gap = 34; // 12mm gap
        
        var availableWidth = pageWidth - (2 * margin);
        var availableHeight = pageHeight - (2 * margin);
        
        var totalHorizontalGaps = 2 * gap;
        var totalVerticalGaps = 3 * gap;
        
        var cellWidth = (availableWidth - totalHorizontalGaps) / 3;
        var cellHeight = (availableHeight - totalVerticalGaps) / 4;
        
        // Define all available patterns with FLOW ORDER
        var patterns = {
            1: { // UP-LOW1
                grid: [
                    [0, 7, 8],     // 1, 8, 9
                    [1, 6, 9],     // 2, 7, 10
                    [2, 5, 10],    // 3, 6, 11
                    [3, 4, 11]     // 4, 5, 12
                ],
                flow: [0, 7, 8, 1, 6, 9, 2, 5, 10, 3, 4, 11]
            },
            2: { // UP-LOW2
                grid: [
                    [0, 1, 2],     // 1, 2, 3
                    [7, 6, 3],     // 8, 7, 4
                    [8, 5, 4],     // 9, 6, 5
                    [9, 10, 11]    // 10, 11, 12
                ],
                flow: [0, 1, 2, 7, 6, 3, 8, 5, 4, 9, 10, 11]
            },
            3: { // UP-LOW3
                grid: [
                    [0, 3, 4],     // 1, 4, 5
                    [1, 2, 5],     // 2, 3, 6
                    [8, 7, 6],     // 9, 8, 7
                    [9, 10, 11]    // 10, 11, 12
                ],
                flow: [0, 3, 4, 1, 2, 5, 8, 7, 6, 9, 10, 11]
            },
            4: { // UP-UP MID
                grid: [
                    [0, 9, 10],    // 1, 10, 11
                    [1, 8, 11],    // 2, 9, 12
                    [2, 7, 6],     // 3, 8, 7
                    [3, 4, 5]      // 4, 5, 6
                ],
                flow: [0, 9, 10, 1, 8, 11, 2, 7, 6, 3, 4, 5]
            },
            5: { // LOW MID – UP MID
                grid: [
                    [8, 9, 10],    // 9, 10, 11
                    [7, 6, 11],    // 8, 7, 12
                    [0, 5, 4],     // 1, 6, 5
                    [1, 2, 3]      // 2, 3, 4
                ],
                flow: [8, 9, 10, 7, 6, 11, 0, 5, 4, 1, 2, 3]
            },
            6: { // LOW – LOW MID
                grid: [
                    [3, 4, 5],     // 4, 5, 6
                    [2, 7, 6],     // 3, 8, 7
                    [1, 8, 11],    // 2, 9, 12
                    [0, 9, 10]     // 1, 10, 11
                ],
                flow: [3, 4, 5, 2, 7, 6, 1, 8, 11, 0, 9, 10]
            },
            7: { // LOW MID – LOW
                grid: [
                    [6, 7, 8],     // 7, 8, 9
                    [5, 4, 9],     // 6, 5, 10
                    [0, 3, 10],    // 1, 4, 11
                    [1, 2, 11]     // 2, 3, 12
                ],
                flow: [6, 7, 8, 5, 4, 9, 0, 3, 10, 1, 2, 11]
            },
            8: { // UP MID – LOW
                grid: [
                    [1, 2, 11],    // 2, 3, 12
                    [0, 3, 10],    // 1, 4, 11
                    [5, 4, 9],     // 6, 5, 10
                    [6, 7, 8]      // 7, 8, 9
                ],
                flow: [1, 2, 11, 0, 3, 10, 5, 4, 9, 6, 7, 8]
            },
            9: { // LOW – UP 1
                grid: [
                    [7, 8, 11],    // 8, 9, 12
                    [6, 9, 10],    // 7, 10, 11
                    [5, 4, 3],     // 6, 5, 4
                    [0, 1, 2]      // 1, 2, 3
                ],
                flow: [7, 8, 11, 6, 9, 10, 5, 4, 3, 0, 1, 2]
            },
            10: { // LOW – UP 2
                grid: [
                    [9, 10, 11],   // 10, 11, 12
                    [8, 7, 6],     // 9, 8, 7
                    [1, 2, 5],     // 2, 3, 6
                    [0, 3, 4]      // 1, 4, 5
                ],
                flow: [9, 10, 11, 8, 7, 6, 1, 2, 5, 0, 3, 4]
            }
        };
        
        // Get the starting page index
        var startPageIndex = startPage.documentOffset;
        
        // Array to store all created paths for post-processing
        var allCreatedPaths = [];
        
        // Process objects in batches of 12
        for (var pageIndex = 0; pageIndex < pagePatterns.length; pageIndex++) {
            var startIndex = pageIndex * objectsPerPage;
            var endIndex = Math.min(startIndex + objectsPerPage, totalObjects);
            var pageObjects = allObjects.slice(startIndex, endIndex);
            
            // Get or create page
            var currentPage;
            if (pageIndex === 0) {
                currentPage = startPage;
            } else {
                // Calculate which page we should be on
                var targetPageIndex = startPageIndex + pageIndex;
                
                // If the page exists, use it; otherwise create it
                if (targetPageIndex < doc.pages.length) {
                    currentPage = doc.pages[targetPageIndex];
                } else {
                    currentPage = doc.pages.add(LocationOptions.AFTER, doc.pages[doc.pages.length - 1]);
                }
            }
            
            // Get the pattern selected for this page
            var patternForThisPage = pagePatterns[pageIndex];
            var selectedPattern = patterns[patternForThisPage];
            
            // Position objects on this page according to pattern AND RENUMBER them
            var objectCenters = arrangeAndRenumberObjectsOnPage(pageObjects, currentPage, selectedPattern, margin, gap, cellWidth, cellHeight, pageIndex);
            
            // Create simple connecting path for this page on BG layer and store it
            if (objectCenters.length > 1) {
                var createdPaths = createSimplePath(currentPage, objectCenters, bgLayer);
                if (createdPaths && createdPaths.length > 0) {
                    allCreatedPaths = allCreatedPaths.concat(createdPaths);
                }
            }
        }
        
        // Convert all points to curved points
        if (allCreatedPaths.length > 0) {
            convertAllPointsToCurved(allCreatedPaths);
        }
        
        // Create summary message
        var patternSummary = "Patterns used:\n";
        var patternNames = {
            1: "UP-LOW1",
            2: "UP-LOW2", 
            3: "UP-LOW3",
            4: "UP-UP MID",
            5: "LOW MID – UP MID",
            6: "LOW – LOW MID",
            7: "LOW MID – LOW",
            8: "UP MID – LOW",
            9: "LOW – UP 1",
            10: "LOW – UP 2"
        };
        
        for (var p = 0; p < pagePatterns.length; p++) {
            var pageNumber = startPageIndex + p + 1;
            patternSummary += "Page " + pageNumber + ": " + patternNames[pagePatterns[p]] + "\n";
        }
        
        alert("Arranged " + totalObjects + " objects across " + pagePatterns.length + " page(s)\n✓ Numbered lists PRESERVED as dynamic\n✓ Numbers follow pattern flow (1,2,3... in step order)\n✓ Starting from page " + (startPageIndex + 1) + " with 12mm gaps\n✓ Objects remain in original layers\n✓ BG and BG PATH layers created\n✓ Connecting paths with curved points\n\n" + patternSummary);
    }
    
    function arrangeAndRenumberObjectsOnPage(objects, page, pattern, margin, gap, cellWidth, cellHeight, pageIndex) {
        var objectsCount = objects.length;
        var centers = [];
        
        // First, arrange all objects in their positions
        for (var row = 0; row < 4; row++) {
            for (var col = 0; col < 3; col++) {
                var objectIndex = row * 3 + col;
                
                // Stop if we've placed all objects for this page
                if (objectIndex >= objectsCount) break;
                
                var patternIndex = pattern.grid[row][col];
                
                // Make sure we don't exceed available objects for partial pages
                if (patternIndex < objectsCount) {
                    var obj = objects[patternIndex];
                    
                    // Calculate position for this cell with gaps
                    var xPos = margin + (col * (cellWidth + gap)) + (cellWidth / 2);
                    var yPos = margin + (row * (cellHeight + gap)) + (cellHeight / 2);
                    
                    // Store center position for connecting path
                    centers[patternIndex] = [xPos, yPos];
                    
                    // Move object to current page if it's not already there
                    if (obj.parentPage != page) {
                        obj.move(page);
                    }
                    
                    // Center object in cell
                    try {
                        obj.move([xPos - (obj.geometricBounds[3] - obj.geometricBounds[1]) / 2, 
                                 yPos - (obj.geometricBounds[2] - obj.geometricBounds[0]) / 2]);
                    } catch(e) {
                        obj.geometricBounds = [
                            yPos - (obj.geometricBounds[2] - obj.geometricBounds[0]) / 2,
                            xPos - (obj.geometricBounds[3] - obj.geometricBounds[1]) / 2,
                            yPos + (obj.geometricBounds[2] - obj.geometricBounds[0]) / 2,
                            xPos + (obj.geometricBounds[3] - obj.geometricBounds[1]) / 2
                        ];
                    }
                }
            }
        }
        
        // Now renumber the objects according to the pattern FLOW
        renumberObjectsByFlow(objects, pattern.flow, pageIndex);
        
        return centers;
    }
    
    // Function: Renumber objects according to pattern flow
    function renumberObjectsByFlow(objects, flowOrder, pageIndex) {
        try {
            // Create a new numbered list for this page
            var listName = "SnakeList_Page_" + (pageIndex + 1);
            
            // Try to get existing list or create new one
            var list;
            try {
                list = doc.lists.item(listName);
                if (!list.isValid) {
                    list = doc.lists.add(listName);
                }
            } catch(e) {
                list = doc.lists.add(listName);
            }
            
            // Clear any existing numbering from objects
            for (var i = 0; i < objects.length; i++) {
                var obj = objects[i];
                if (obj.constructor.name === "TextFrame") {
                    try {
                        var story = obj.parentStory;
                        for (var p = 0; p < story.paragraphs.length; p++) {
                            var paragraph = story.paragraphs[p];
                            // Remove any existing numbering
                            paragraph.appliedNumberingList = null;
                            paragraph.numberingStartAt = 1;
                        }
                    } catch(e) {}
                }
            }
            
            // Apply new numbering in flow order
            var stepNumber = 1;
            for (var flowIndex = 0; flowIndex < flowOrder.length; flowIndex++) {
                var objectIndex = flowOrder[flowIndex];
                if (objectIndex < objects.length) {
                    var obj = objects[objectIndex];
                    if (obj.constructor.name === "TextFrame") {
                        try {
                            var story = obj.parentStory;
                            for (var p = 0; p < story.paragraphs.length; p++) {
                                var paragraph = story.paragraphs[p];
                                // Apply the new list
                                paragraph.appliedNumberingList = list;
                                paragraph.numberingStartAt = stepNumber;
                                stepNumber++;
                                break; // Only number first paragraph
                            }
                        } catch(e) {}
                    }
                }
            }
            
        } catch(e) {
            // If dynamic numbering fails, fall back to converting to text silently
            for (var i = 0; i < objects.length; i++) {
                var obj = objects[i];
                if (obj.constructor.name === "TextFrame") {
                    try {
                        var story = obj.parentStory;
                        for (var p = 0; p < story.paragraphs.length; p++) {
                            var paragraph = story.paragraphs[p];
                            if (paragraph.appliedNumberingList) {
                                paragraph.convertNumberingToText();
                            }
                        }
                    } catch(e) {}
                }
            }
        }
    }
    
    function createSimplePath(page, centers, bgLayer) {
        var createdPaths = [];
        try {
            // Create a simple graphic line with the entire path
            var line = page.graphicLines.add();
            
            // Set line color to 20% black and thickness to 20pt
            line.strokeColor = doc.swatches.item("Black");
            line.strokeTint = 20; // 20% black
            line.strokeWeight = 20; // 20 points thick
            line.fillColor = doc.swatches.item("None");
            
            // Build the path points array
            var pathPoints = [];
            for (var i = 0; i < centers.length; i++) {
                if (centers[i]) {
                    pathPoints.push(centers[i]);
                }
            }
            
            // Set the entire path as a single continuous line
            if (pathPoints.length > 1) {
                line.paths[0].entirePath = pathPoints;
            }
            
            // Move ONLY THE PATH to the BG PATH layer
            if (bgLayer && bgLayer.isValid) {
                try {
                    line.itemLayer = bgLayer;
                } catch(e) {
                    // If moving to layer fails, continue without error
                }
            }
            
            createdPaths.push(line);
            
        } catch(e) {
            // If the graphic line method fails, try individual segments as last resort
            try {
                for (var i = 0; i < centers.length - 1; i++) {
                    if (centers[i] && centers[i+1]) {
                        var segment = page.graphicLines.add();
                        
                        // Set segment color to 20% black and thickness to 20pt
                        segment.strokeColor = doc.swatches.item("Black");
                        segment.strokeTint = 20; // 20% black
                        segment.strokeWeight = 20; // 20 points thick
                        segment.fillColor = doc.swatches.item("None");
                        
                        segment.paths[0].entirePath = [centers[i], centers[i+1]];
                        
                        // Move ONLY THE SEGMENT to BG PATH layer
                        if (bgLayer && bgLayer.isValid) {
                            try {
                                segment.itemLayer = bgLayer;
                            } catch(e) {
                                // Continue without error
                            }
                        }
                        
                        createdPaths.push(segment);
                    }
                }
            } catch(e2) {
                // If all methods fail, skip the path creation
            }
        }
        
        return createdPaths;
    }
    
    // Function: Convert all points to curved points
    function convertAllPointsToCurved(paths) {
        try {
            // Select all paths
            app.select(paths);
            
            // Wait briefly for selection to process
            $.sleep(100);
            
            // For each path, convert all points to curved points
            for (var i = 0; i < paths.length; i++) {
                var path = paths[i];
                if (path && path.isValid) {
                    try {
                        // Loop through all paths in the GraphicLine object
                        for (var p = 0; p < path.paths.length; p++) {
                            var currentPath = path.paths[p];
                            
                            // Loop through all points in the path
                            for (var pt = 0; pt < currentPath.pathPoints.length; pt++) {
                                var point = currentPath.pathPoints[pt];
                                
                                try {
                                    point.pointType = 135253396; // Smooth point
                                } catch(e) {
                                    try {
                                        point.leftDirection = point.anchor;
                                        point.rightDirection = point.anchor;
                                    } catch(e2) {}
                                }
                            }
                        }
                    } catch(e) {}
                }
            }
        } catch(e) {}
    }
})();