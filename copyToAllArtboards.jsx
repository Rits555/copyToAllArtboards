#target illustrator

(function () {
    var doc = app.activeDocument;
    if (!doc) {
        alert("No active document found.");
        return;
    }

    if (!doc.selection || doc.selection.length === 0) {
        alert("Please select at least one object.");
        return;
    }

    var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();

    var dialog = new Window("dialog", "Copy to Artboards");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.spacing = 12;
    dialog.margins = 16;

    // Target Selection Group
    var targetGroup = dialog.add("panel", undefined, "Target");
    targetGroup.orientation = "column";
    targetGroup.alignChildren = "left";
    targetGroup.spacing = 8;
    
    var allBtn = targetGroup.add("radiobutton", undefined, "Copy to all artboards");
    allBtn.value = true;
    
    var selectedBtn = targetGroup.add("radiobutton", undefined, "Copy to selected artboards");
    
    var selectedPanel = targetGroup.add("group");
    selectedPanel.orientation = "column";
    selectedPanel.alignChildren = "fill";
    selectedPanel.margins = [20, 0, 0, 0];
    var artboardLabel = selectedPanel.add("statictext", undefined, "Enter artboard indices (comma-separated):");
    var artboardInput = selectedPanel.add("edittext", undefined, "1,2,3");
    artboardInput.characters = 30;
    artboardInput.enabled = false;

    // Placement Mode Group
    var placementGroup = dialog.add("panel", undefined, "Placement Mode");
    placementGroup.orientation = "column";
    placementGroup.alignChildren = "left";
    placementGroup.spacing = 8;
    
    var samePositionBtn = placementGroup.add("radiobutton", undefined, "Same position");
    samePositionBtn.value = true;
    
    var samePropBtn = placementGroup.add("radiobutton", undefined, "Resize in same proportion");
    var samePropDesc = placementGroup.add("statictext", undefined, "");
    samePropDesc.text = "(scale relative to target artboard size)";
    samePropDesc.margins = [20, 0, 0, 0];
    samePropDesc.graphics.font = ScriptUI.newFont("helvetica", "condensed", 10);
    
    var centeredBtn = placementGroup.add("radiobutton", undefined, "Centered on artboard");

    // Options Group
    var optionsGroup = dialog.add("panel", undefined, "Options");
    optionsGroup.orientation = "column";
    optionsGroup.alignChildren = "left";
    optionsGroup.spacing = 8;
    
    var skipCurrentCheckbox = optionsGroup.add("checkbox", undefined, "Skip current artboard");
    skipCurrentCheckbox.value = false;

    // Button Group
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignChildren = "center";
    buttonGroup.spacing = 10;
    
    var okBtn = buttonGroup.add("button", undefined, "Copy", { name: "ok" });
    var cancelBtn = buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });

    // Event handlers
    allBtn.onClick = function () {
        artboardInput.enabled = false;
        artboardLabel.enabled = false;
    };

    selectedBtn.onClick = function () {
        artboardInput.enabled = true;
        artboardLabel.enabled = true;
        artboardInput.active = true;
    };

    okBtn.onClick = function () {
        var targetMode = allBtn.value ? "all" : "selected";
        var placementMode = "same_position";
        if (samePropBtn.value) {
            placementMode = "same_proportion";
        } else if (centeredBtn.value) {
            placementMode = "centered";
        }
        var skipCurrent = skipCurrentCheckbox.value;

        // Get target artboards
        var artboardIndexes = [];
        if (targetMode === "all") {
            for (var i = 0; i < doc.artboards.length; i++) {
                artboardIndexes.push(i);
            }
        } else {
            var raw = artboardInput.text;
            var parts = raw.split(",");
            for (var i = 0; i < parts.length; i++) {
                var trimmed = parts[i].replace(/^\s+|\s+$/g, "");
                var n = parseInt(trimmed, 10);
                if (!isNaN(n) && n > 0 && n <= doc.artboards.length) {
                    artboardIndexes.push(n - 1); // Convert to 0-based index
                }
            }
            if (artboardIndexes.length === 0) {
                alert("No valid artboard indices entered.\nExample: 1,2,3 (artboards are numbered starting from 1)");
                return;
            }
        }

        // Remove current artboard if skip is enabled
        if (skipCurrent) {
            artboardIndexes = artboardIndexes.filter(function (idx) {
                return idx !== activeArtboardIndex;
            });
        }

        if (artboardIndexes.length === 0) {
            alert("No artboards selected. All artboards were skipped or list was empty.");
            return;
        }

        // Store original selection
        var originalSelection = [];
        for (var i = 0; i < doc.selection.length; i++) {
            originalSelection.push(doc.selection[i]);
        }

        var currentArtboard = doc.artboards[activeArtboardIndex];
        var currentRect = getArtboardRect(currentArtboard);

        // Process each target artboard
        for (var a = 0; a < artboardIndexes.length; a++) {
            var artboardIndex = artboardIndexes[a];
            if (artboardIndex < 0 || artboardIndex >= doc.artboards.length) continue;

            doc.artboards.setActiveArtboardIndex(artboardIndex);

            var targetArtboard = doc.artboards[artboardIndex];
            var targetRect = getArtboardRect(targetArtboard);

            // Duplicate items to this artboard
            for (var s = 0; s < originalSelection.length; s++) {
                var item = originalSelection[s];
                if (item && item.duplicate) {
                    var clone = item.duplicate();

                    // Apply placement mode
                    switch (placementMode) {
                        case "same_position":
                            placeSamePosition(clone, currentRect, targetRect);
                            break;

                        case "same_proportion":
                            placeSameProportion(clone, currentRect, targetRect);
                            break;

                        case "centered":
                            centerOnArtboard(clone, targetRect);
                            break;
                    }
                }
            }
        }

        // Return to original artboard
        doc.artboards.setActiveArtboardIndex(activeArtboardIndex);
        dialog.close();
        alert("Successfully copied to " + artboardIndexes.length + " artboard(s).");
    };

    cancelBtn.onClick = function () {
        dialog.close();
    };

    dialog.show();

    // Helper functions
    function getArtboardRect(artboard) {
        var rect = artboard.artboardRect;
        return {
            left: rect[0],
            top: rect[1],
            right: rect[2],
            bottom: rect[3],
            width: Math.abs(rect[2] - rect[0]),
            height: Math.abs(rect[1] - rect[3])
        };
    }

    function placeSamePosition(item, currentRect, targetRect) {
        var bounds = item.visibleBounds;
        if (bounds[0] === bounds[2] || bounds[1] === bounds[3]) return; // Skip if no bounds

        var itemLeft = bounds[0];
        var itemTop = bounds[1];

        // Calculate relative position within current artboard
        var localX = itemLeft - currentRect.left;
        var localY = itemTop - currentRect.top;

        // Apply same relative position to target artboard
        item.left = targetRect.left + localX;
        item.top = targetRect.top + localY;
    }

    function placeSameProportion(item, currentRect, targetRect) {
        var bounds = item.visibleBounds;
        if (bounds[0] === bounds[2] || bounds[1] === bounds[3]) return; // Skip if no bounds

        var itemW = Math.abs(bounds[2] - bounds[0]);
        var itemH = Math.abs(bounds[1] - bounds[3]);
        var itemLeft = bounds[0];
        var itemTop = bounds[1];

        // Calculate relative position and size within current artboard
        var relativeX = (itemLeft - currentRect.left) / currentRect.width;
        var relativeY = (itemTop - currentRect.top) / currentRect.height;
        var relativeW = itemW / currentRect.width;
        var relativeH = itemH / currentRect.height;

        // Apply to target artboard with same proportions
        var newLeft = targetRect.left + (relativeX * targetRect.width);
        var newTop = targetRect.top + (relativeY * targetRect.height);
        var newW = relativeW * targetRect.width;
        var newH = relativeH * targetRect.height;

        // Set new dimensions
        var scaleX = newW / itemW * 100;
        var scaleY = newH / itemH * 100;

        item.resize(scaleX, scaleY, true, true, true, true, 1, Transformation.DOCUMENTORIGIN);

        // Set new position (after resize, adjust for center point)
        item.left = newLeft + (newW / 2) - (itemW * (scaleX / 100) / 2);
        item.top = newTop + (newH / 2) - (itemH * (scaleY / 100) / 2);
    }

    function centerOnArtboard(item, targetRect) {
        var bounds = item.visibleBounds;
        if (bounds[0] === bounds[2] || bounds[1] === bounds[3]) return; // Skip if no bounds

        var itemCenterX = (bounds[0] + bounds[2]) / 2;
        var itemCenterY = (bounds[1] + bounds[3]) / 2;

        var artboardCenterX = targetRect.left + (targetRect.width / 2);
        var artboardCenterY = targetRect.top + (targetRect.height / 2);

        var dx = artboardCenterX - itemCenterX;
        var dy = artboardCenterY - itemCenterY;

        item.translate(dx, dy);
    }
})();