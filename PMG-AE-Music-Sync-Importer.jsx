/*
  PainnMyGraphic XML Pipeline
  ------------------------------------------------------------
  v3.1
  功能：
  - 讀取 Premiere/FCP XML（xmeml v4 類型）
  - 建立 AE Comp
  - 依照 XML 中的 audio clipitem 自動匯入音檔並排入時間軸
  - 自動依檔名分類：BGM / AMB / SFX / OTHER
  - 自動套用 layerLabel 顏色
  - 自動建立 comp marker（每個 clip 的開始點）
  - 自動建立說明文字層 README
  - 自動建立品牌角標 PainnMyGraphic

  使用方式：
  1. 在 After Effects 執行 File > Scripts > Run Script File...
  2. 選這支 jsx
  3. 選擇 XML 檔
  4. 選擇音檔資料夾
  5. 腳本會建立 comp 並自動排入音訊、marker、README、品牌角標
*/

(function PainnMyGraphic_XML_Pipeline() {
    app.beginUndoGroup("PainnMyGraphic XML Pipeline");

    function stopWithAlert(msg) {
        alert(msg);
        throw new Error(msg);
    }

    function trim(s) {
        return String(s || "").replace(/^\s+|\s+$/g, "");
    }

    function toInt(v, fallback) {
        var n = parseInt(v, 10);
        return isNaN(n) ? fallback : n;
    }

    function sanitizeName(name) {
        if (!name) return "Untitled";
        return String(name).replace(/[\\\/\:\*\?\"\<\>\|]/g, "_");
    }

    function decodePathUrl(pathurl) {
        if (!pathurl) return "";
        var s = String(pathurl);
        s = s.replace(/^file:\/\//i, "");
        s = s.replace(/%20/g, " ")
             .replace(/%2C/gi, ",")
             .replace(/%26/gi, "&")
             .replace(/%27/gi, "'")
             .replace(/%28/gi, "(")
             .replace(/%29/gi, ")")
             .replace(/%5B/gi, "[")
             .replace(/%5D/gi, "]")
             .replace(/%23/gi, "#")
             .replace(/%40/gi, "@")
             .replace(/%2B/gi, "+")
             .replace(/%3D/gi, "=")
             .replace(/%3B/gi, ";")
             .replace(/%3A/gi, ":")
             .replace(/%21/gi, "!")
             .replace(/%24/gi, "$")
             .replace(/%5E/gi, "^")
             .replace(/%60/gi, "`")
             .replace(/%7E/gi, "~");
        return s;
    }

    function safeText(xmlNode, childName, defaultValue) {
        try {
            if (childName) {
                var node = xmlNode[childName][0];
                if (node != undefined && node.toString() !== "") {
                    return trim(node.toString());
                }
            } else {
                if (xmlNode != undefined && xmlNode.toString() !== "") {
                    return trim(xmlNode.toString());
                }
            }
        } catch (e) {}
        return defaultValue;
    }

    function safeAttr(xmlNode, attrName, defaultValue) {
        try {
            var a = xmlNode.attribute(attrName);
            if (a != undefined && a.toString() !== "") {
                return a.toString();
            }
        } catch (e) {}
        return defaultValue;
    }

    function ensureProject() {
        if (!app.project) {
            app.newProject();
        }
    }

    function buildFileMap(folder) {
        var map = {};

        function walk(f) {
            var items = f.getFiles();
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (item instanceof Folder) {
                    walk(item);
                } else if (item instanceof File) {
                    map[item.name.toLowerCase()] = item;
                }
            }
        }

        if (folder && folder.exists) {
            walk(folder);
        }
        return map;
    }

    function fileNameFromPath(path) {
        if (!path) return "";
        var s = String(path).replace(/\\/g, "/");
        var parts = s.split("/");
        return parts.length ? parts[parts.length - 1] : s;
    }

    function importFootageOnce(fileObj, cache) {
        var key = fileObj.fsName;
        if (cache[key]) {
            return cache[key];
        }

        var io = new ImportOptions(fileObj);
        var imported = app.project.importFile(io);
        cache[key] = imported;
        return imported;
    }

    function resolveMediaFile(pathUrl, fileName, chosenFolder, fileMap) {
        var decodedPath = decodePathUrl(pathUrl);

        if (decodedPath) {
            try {
                var directFile = new File(decodedPath);
                if (directFile.exists) {
                    return directFile;
                }
            } catch (e1) {}
        }

        if (fileName) {
            var lower = fileName.toLowerCase();
            if (fileMap[lower] && fileMap[lower].exists) {
                return fileMap[lower];
            }
        }

        if (decodedPath) {
            var fromPathName = fileNameFromPath(decodedPath);
            var lower2 = fromPathName.toLowerCase();
            if (fileMap[lower2] && fileMap[lower2].exists) {
                return fileMap[lower2];
            }
        }

        if (fileName && chosenFolder && chosenFolder.exists) {
            var fallback = new File(chosenFolder.fsName + "/" + fileName);
            if (fallback.exists) {
                return fallback;
            }
        }

        return null;
    }

    function collectFileDefinitions(xmlRoot) {
        var map = {};
        var allFiles = xmlRoot..file;

        for each (var f in allFiles) {
            var id = safeAttr(f, "id", "");
            if (!id) continue;

            var existing = map[id] || {
                id: id,
                name: "",
                pathurl: ""
            };

            var nm = safeText(f, "name", "");
            var pu = safeText(f, "pathurl", "");

            if (nm) existing.name = nm;
            if (pu) existing.pathurl = pu;

            map[id] = existing;
        }

        return map;
    }

    function classifyClipName(name) {
        var s = String(name || "").toLowerCase();

        if (
            s.indexOf("states") !== -1 ||
            s.indexOf("score") !== -1 ||
            s.indexOf("music") !== -1 ||
            s.indexOf("theme") !== -1 ||
            s.indexOf("song") !== -1
        ) {
            return "BGM";
        }

        if (
            s.indexOf("ambience") !== -1 ||
            s.indexOf("room tone") !== -1 ||
            s.indexOf("drone") !== -1 ||
            s.indexOf("forest") !== -1 ||
            s.indexOf("jungle") !== -1 ||
            s.indexOf("museum") !== -1 ||
            s.indexOf("office") !== -1 ||
            s.indexOf("hallway") !== -1 ||
            s.indexOf("lab") !== -1 ||
            s.indexOf("laboratory") !== -1 ||
            s.indexOf("scifi") !== -1 ||
            s.indexOf("science") !== -1 ||
            s.indexOf("control room") !== -1
        ) {
            return "AMB";
        }

        if (
            s.indexOf("footsteps") !== -1 ||
            s.indexOf("water") !== -1 ||
            s.indexOf("drip") !== -1 ||
            s.indexOf("click") !== -1 ||
            s.indexOf("button") !== -1 ||
            s.indexOf("mechanical") !== -1 ||
            s.indexOf("wings") !== -1 ||
            s.indexOf("door") !== -1 ||
            s.indexOf("bubbles") !== -1 ||
            s.indexOf("growl") !== -1 ||
            s.indexOf("vocalisation") !== -1 ||
            s.indexOf("zap") !== -1 ||
            s.indexOf("handling") !== -1 ||
            s.indexOf("cloth") !== -1 ||
            s.indexOf("creature") !== -1 ||
            s.indexOf("insect") !== -1 ||
            s.indexOf("electricity") !== -1
        ) {
            return "SFX";
        }

        return "OTHER";
    }

    function getLabelIndexByCategory(cat) {
        if (cat === "BGM") return 9;
        if (cat === "AMB") return 8;
        if (cat === "SFX") return 2;
        return 11;
    }

    function parseSequence(xmlRoot, fileDefMap) {
        var sequence = xmlRoot..sequence[0];
        if (sequence == undefined || sequence.toString() === "") {
            stopWithAlert(
                "找不到 <sequence>，請確認 XML 格式是否正確。\n\n" +
                "若是 Premiere 匯出的 XML，一般根節點應該是 <xmeml>。\n" +
                "若是 FCPXML，格式不同，需要另一版解析器。"
            );
        }

        var sequenceName = safeText(sequence, "name", "XML_Imported_Timeline");
        var durationFrames = toInt(safeText(sequence, "duration", "0"), 0);
        var fps = 24;

        try {
            var rateNode = sequence.rate[0];
            if (rateNode != undefined) {
                fps = toInt(safeText(rateNode, "timebase", "24"), 24);
            }
        } catch (e) {}

        if (!fps || fps <= 0) fps = 24;

        var audioNode = sequence.media[0].audio[0];
        if (audioNode == undefined || audioNode.toString() === "") {
            stopWithAlert("目前腳本只支援 audio 節點，但 XML 裡找不到 audio。");
        }

        var tracks = [];
        var trackNodes = audioNode.track;

        for (var t = 0; t < trackNodes.length(); t++) {
            var track = trackNodes[t];
            var clips = [];
            var clipNodes = track.clipitem;

            for (var c = 0; c < clipNodes.length(); c++) {
                var clip = clipNodes[c];
                var clipId = safeAttr(clip, "id", "");
                var clipName = safeText(clip, "name", "");
                var start = toInt(safeText(clip, "start", "0"), 0);
                var end = toInt(safeText(clip, "end", "0"), 0);
                var inPoint = toInt(safeText(clip, "in", "0"), 0);
                var outPoint = toInt(safeText(clip, "out", "0"), 0);

                var fileNode = clip.file[0];
                var fileRefId = "";
                var fileName = "";
                var pathUrl = "";

                if (fileNode != undefined) {
                    fileRefId = safeAttr(fileNode, "id", "");
                    fileName = safeText(fileNode, "name", "");
                    pathUrl = safeText(fileNode, "pathurl", "");

                    if (fileRefId && fileDefMap[fileRefId]) {
                        if (!fileName) fileName = fileDefMap[fileRefId].name;
                        if (!pathUrl) pathUrl = fileDefMap[fileRefId].pathurl;
                    }
                }

                if (!fileName && clipName) {
                    fileName = clipName;
                }

                if (!clipName && !fileName && start === 0 && end === 0 && inPoint === 0 && outPoint === 0) {
                    continue;
                }

                var displayName = clipName || fileName || "Unnamed";
                var category = classifyClipName(displayName);

                clips.push({
                    clipId: clipId,
                    clipName: clipName,
                    start: start,
                    end: end,
                    inPoint: inPoint,
                    outPoint: outPoint,
                    fileRefId: fileRefId,
                    fileName: fileName,
                    pathUrl: pathUrl,
                    trackIndex: t + 1,
                    category: category,
                    displayName: displayName
                });
            }

            tracks.push({
                trackIndex: t + 1,
                clips: clips
            });
        }

        return {
            name: sequenceName,
            fps: fps,
            durationFrames: durationFrames,
            tracks: tracks
        };
    }

    function addCompMarker(comp, timeSec, labelText, commentText) {
        try {
            var mv = new MarkerValue(labelText);
            mv.comment = commentText || labelText;
            comp.markerProperty.setValueAtTime(timeSec, mv);
        } catch (e) {}
    }

    function createReadmeLayer(comp, summaryText) {
        try {
            var textLayer = comp.layers.addText(summaryText);
            textLayer.name = "README_XML_IMPORT";
            textLayer.label = 14;
            textLayer.startTime = 0;
            textLayer.inPoint = 0;
            textLayer.outPoint = Math.min(comp.duration, 20);

            var textProp = textLayer.property("Source Text");
            var textDoc = textProp.value;
            textDoc.fontSize = 42;
            textDoc.leading = 52;
            textDoc.fillColor = [1, 1, 1];
            textDoc.applyFill = true;
            textDoc.applyStroke = false;
            textDoc.justification = ParagraphJustification.LEFT_JUSTIFY;
            textProp.setValue(textDoc);

            textLayer.property("Position").setValue([260, 180]);
            return textLayer;
        } catch (e) {
            return null;
        }
    }

    function createBrandLayer(comp) {
        try {
            var brandLayer = comp.layers.addText("PainnMyGraphic");
            brandLayer.name = "BRAND_PainnMyGraphic";
            brandLayer.label = 10;
            brandLayer.startTime = 0;
            brandLayer.inPoint = 0;
            brandLayer.outPoint = comp.duration;

            var brandTextProp = brandLayer.property("Source Text");
            var brandTextDoc = brandTextProp.value;
            brandTextDoc.fontSize = 48;
            brandTextDoc.fillColor = [1, 1, 1];
            brandTextDoc.applyFill = true;
            brandTextDoc.applyStroke = false;
            brandTextDoc.justification = ParagraphJustification.RIGHT_JUSTIFY;
            brandTextProp.setValue(brandTextDoc);

            brandLayer.property("Position").setValue([comp.width - 200, 120]);
            return brandLayer;
        } catch (e) {
            return null;
        }
    }

    ensureProject();

    var xmlFile = File.openDialog("選擇 Premiere / FCP XML 檔", "XML:*.xml,All Files:*.*", false);
    if (!xmlFile) {
        app.endUndoGroup();
        return;
    }

    var audioFolder = Folder.selectDialog("選擇音檔資料夾（可包含子資料夾）");
    if (!audioFolder) {
        app.endUndoGroup();
        return;
    }

    if (!xmlFile.exists) {
        stopWithAlert("XML 檔不存在。\n" + xmlFile.fsName);
    }
    if (!audioFolder.exists) {
        stopWithAlert("音檔資料夾不存在。\n" + audioFolder.fsName);
    }

    xmlFile.encoding = "UTF-8";
    if (!xmlFile.open("r")) {
        stopWithAlert("無法開啟 XML 檔。\n" + xmlFile.fsName);
    }

    var xmlString = xmlFile.read();
    xmlFile.close();

    var xmlRoot;
    try {
        xmlRoot = new XML(xmlString);
    } catch (e) {
        stopWithAlert("XML 解析失敗，請確認格式是否正確。\n\n" + e.toString());
    }

    var fileDefMap = collectFileDefinitions(xmlRoot);
    var sequenceData = parseSequence(xmlRoot, fileDefMap);

    var fps = sequenceData.fps;
    var compDuration = sequenceData.durationFrames / fps;
    if (compDuration <= 0) compDuration = 10;

    var compName = sanitizeName(sequenceData.name || "XML_Imported_Timeline");
    var comp = app.project.items.addComp(compName, 1920, 1080, 1, compDuration, fps);

    var rootFolder = app.project.items.addFolder(compName + "_Assets");
    comp.parentFolder = rootFolder;

    var audioBin = app.project.items.addFolder("Audio");
    audioBin.parentFolder = rootFolder;

    var importedCache = {};
    var fileMap = buildFileMap(audioFolder);
    var missingFiles = [];
    var placedCount = 0;
    var categoryCounts = {
        BGM: 0,
        AMB: 0,
        SFX: 0,
        OTHER: 0
    };

    for (var t = sequenceData.tracks.length - 1; t >= 0; t--) {
        var track = sequenceData.tracks[t];

        for (var c = 0; c < track.clips.length; c++) {
            var clip = track.clips[c];
            var resolved = resolveMediaFile(clip.pathUrl, clip.fileName, audioFolder, fileMap);
            var startTime = clip.start / fps;
            var layerDuration = (clip.end - clip.start) / fps;

            if (layerDuration < 0) layerDuration = 0;

            addCompMarker(
                comp,
                startTime,
                clip.category + " | T" + clip.trackIndex,
                clip.displayName + "\nTrack: " + clip.trackIndex + "\nStart: " + startTime.toFixed(2) + "s"
            );

            if (!resolved) {
                missingFiles.push(
                    "Track " + clip.trackIndex +
                    " | [" + clip.category + "] " + clip.displayName +
                    (clip.pathUrl ? "\n  pathurl: " + decodePathUrl(clip.pathUrl) : "")
                );
                continue;
            }

            var footage;
            try {
                footage = importFootageOnce(resolved, importedCache);
                footage.parentFolder = audioBin;
            } catch (impErr) {
                missingFiles.push("匯入失敗：" + resolved.fsName + "\n  " + impErr.toString());
                continue;
            }

            var layer;
            try {
                layer = comp.layers.add(footage);
            } catch (layerErr) {
                missingFiles.push("建立圖層失敗：" + resolved.fsName + "\n  " + layerErr.toString());
                continue;
            }

            var inTime = clip.inPoint / fps;
            layer.name = "[" + clip.category + "] T" + clip.trackIndex + " | " + clip.displayName;
            layer.startTime = startTime - inTime;
            layer.inPoint = startTime;
            layer.outPoint = startTime + layerDuration;
            layer.label = getLabelIndexByCategory(clip.category);

            placedCount++;
            categoryCounts[clip.category]++;
        }
    }

    var readme = [];
    readme.push("PainnMyGraphic XML Pipeline");
    readme.push("Comp: " + compName);
    readme.push("FPS: " + fps);
    readme.push("Duration: " + sequenceData.durationFrames + " frames / " + compDuration.toFixed(2) + " sec");
    readme.push("Tracks: " + sequenceData.tracks.length);
    readme.push("Placed: " + placedCount);
    readme.push("BGM: " + categoryCounts.BGM + " | AMB: " + categoryCounts.AMB + " | SFX: " + categoryCounts.SFX + " | OTHER: " + categoryCounts.OTHER);
    readme.push("");
    readme.push("Color Guide");
    readme.push("BGM = Green-ish");
    readme.push("AMB = Blue-ish");
    readme.push("SFX = Red-ish");
    readme.push("OTHER = Purple-ish");
    readme.push("");
    readme.push("Tip: 選 layer 後按 LL 可看音訊波形");
    readme.push("");
    readme.push("PainnMyGraphic");
    readme.push("Website: https://painnmygraphic.com/");
    readme.push("YouTube: https://www.youtube.com/@painnmygraphic");

    if (missingFiles.length) {
        readme.push("");
        readme.push("Missing Files: " + missingFiles.length);

        var maxList = Math.min(missingFiles.length, 8);
        for (var mi = 0; mi < maxList; mi++) {
            readme.push("- " + missingFiles[mi].split("\n")[0]);
        }

        if (missingFiles.length > maxList) {
            readme.push("- ...and more");
        }
    }

    createReadmeLayer(comp, readme.join("\r"));
    createBrandLayer(comp);

    var summary = [];
    summary.push("完成。");
    summary.push("");
    summary.push("Comp: " + compName);
    summary.push("FPS: " + fps);
    summary.push("Duration: " + sequenceData.durationFrames + " frames (" + compDuration.toFixed(2) + " sec)");
    summary.push("Tracks: " + sequenceData.tracks.length);
    summary.push("Placed clips: " + placedCount);
    summary.push("BGM: " + categoryCounts.BGM + " | AMB: " + categoryCounts.AMB + " | SFX: " + categoryCounts.SFX + " | OTHER: " + categoryCounts.OTHER);
    summary.push("Markers: 已建立");
    summary.push("README Layer: 已建立");
    summary.push("Brand Layer: 已建立");

    if (missingFiles.length) {
        summary.push("");
        summary.push("找不到或匯入失敗的檔案：" + missingFiles.length);
    }

    alert(summary.join("\n"));
    app.endUndoGroup();
})();
