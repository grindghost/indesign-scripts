#target "InDesign"

(function () {
  if (!app.documents.length) {
    alert("No open document.");
    return;
  }

  // 🔧 Set units to Picas
  app.scriptPreferences.measurementUnit = MeasurementUnits.PICAS;
  app.documents[0].viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PICAS;
  app.documents[0].viewPreferences.verticalMeasurementUnits = MeasurementUnits.PICAS;

  var doc = app.activeDocument;
  var cornerPrototype = null;
  var nmbPrototype = null;

  // 🔎 Constants
  var TOP_MARGIN_FROM_PAGE = 12; // starting Y position on each page
  var UNIT_TOP_SPACING = 2; // espace entre le bas du rectangle rouge précédent et le haut du unit suivant

  var UNIT_LABEL = "unit";
  var FIELD_PADDING = 0.8;
  var SPACING_BELOW_UNIT = 0;

  var SESSION_ID_LABEL = "formation_id";
  var SESSION_ID = null;
  var METADATA_LAYER_NAME = "_metadatas";

  var metadataLayer = null;
  try {
    metadataLayer = doc.layers.itemByName(METADATA_LAYER_NAME);
    metadataLayer.name;
  } catch (_) {
    metadataLayer = doc.layers.add({ name: METADATA_LAYER_NAME });
  }

  var metaFrame = null;
  var metadata = {};

  // === 🧠 Robust JSON helper (for ExtendScript, supports nested objects/arrays) ===
  function stringifyJSON(obj) {
    if (obj === null) return 'null';
    if (typeof obj === 'string') return '"' + obj.replace(/"/g, '\\"') + '"';
    if (typeof obj === 'number' || typeof obj === 'boolean') return String(obj);
    if (obj instanceof Array) {
      var arr = [];
      for (var i = 0; i < obj.length; i++) arr.push(stringifyJSON(obj[i]));
      return '[' + arr.join(',') + ']';
    }
    var str = '{';
    var first = true;
    for (var key in obj) {
      if (obj.hasOwnProperty(key)) {
        if (!first) str += ',';
        str += '"' + key + '":' + stringifyJSON(obj[key]);
        first = false;
      }
    }
    str += '}';
    return str;
  }

  function parseJSON(str) {
    try {
      return eval('(' + str + ')');
    } catch (e) {
      alert("⚠️ Erreur de lecture des métadonnées JSON.");
      return {};
    }
  }

  function generateRandomString(length) {
    var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    var result = "";
    for (var i = 0; i < length; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  }

  function createOrGetColor(name, rgb) {
    try {
      var c = doc.colors.item(name);
      c.name;
      return c;
    } catch (_) {
      return doc.colors.add({
        name: name,
        model: ColorModel.process,
        space: ColorSpace.RGB,
        colorValue: rgb
      });
    }
  }

  function findMetaFrame() {
    for (var i = 0; i < doc.textFrames.length; i++) {
      var tf = doc.textFrames[i];
      if (tf.label === SESSION_ID_LABEL) return tf;
    }
    return null;
  }

  function writeMetadata(metaObj) {
    if (!metaFrame) return;
    metaFrame.contents = stringifyJSON(metaObj);
  }

  function readMetadata() {
    if (!metaFrame) return {};
    return parseJSON(metaFrame.contents);
  }

  function cleanUnits(metaObj) {
    if (!metaObj.units) return;
    for (var unitId in metaObj.units) {
      var unit = metaObj.units[unitId];
      // Remove rectangle
      if (unit.rectangleId) {
        try {
          var rect = findPageItemByLabel(unit.rectangleId);
          if (rect) rect.remove();
        } catch (e) {}
      }
      // Remove textbox
      if (unit.texBoxId) {
        try {
          var txtb = findPageItemByLabel(unit.texBoxId);
          if (txtb) txtb.remove();
        } catch (e) {}
      }
      // Remove nbm
      if (unit.nbmId) {
        try {
          var nbm = findPageItemByLabel(unit.nbmId);
          if (nbm) nbm.remove();
        } catch (e) {}
      }
      // Remove corner
      if (unit.cornerId) {
        try {
          var corner = findPageItemByLabel(unit.cornerId);
          if (corner) corner.remove();
        } catch (e) {}
      }
    }
    delete metaObj.units;
    writeMetadata(metaObj);
  }

  function findPageItemByLabel(label) {
    var items = doc.allPageItems;
    for (var i = 0; i < items.length; i++) {
      if (items[i].label === label) return items[i];
    }
    return null;
  }

  function updateFormationId(metaObj, newId) {
    if (!metaObj.units) return;
    for (var unitId in metaObj.units) {
      var unit = metaObj.units[unitId];
      // Update rectangle label
      if (unit.rectangleId) {
        var rect = findPageItemByLabel(unit.rectangleId);
        if (rect) rect.label = 'rect_' + newId + '_' + unit.index;
        unit.rectangleId = 'rect_' + newId + '_' + unit.index;
      }
      // Update textbox label and name
      if (unit.texBoxId) {
        var txtb = findPageItemByLabel(unit.texBoxId);
        if (txtb) {
          txtb.label = 'txtb_' + newId + '_' + unit.index;
          txtb.name = newId + '_' + unit.index;
        }
        unit.texBoxId = 'txtb_' + newId + '_' + unit.index;
      }
      // Update nbm label
      if (unit.nbmId) {
        var nbm = findPageItemByLabel(unit.nbmId);
        if (nbm) nbm.label = 'nbm_' + newId + '_' + unit.index;
        unit.nbmId = 'nbm_' + newId + '_' + unit.index;
      }
      // Update corner label
      if (unit.cornerId) {
        var corner = findPageItemByLabel(unit.cornerId);
        if (corner) corner.label = 'corner_' + newId + '_' + unit.index;
        unit.cornerId = 'corner_' + newId + '_' + unit.index;
      }
    }
    metaObj.formation_id = newId;
    writeMetadata(metaObj);
  }

  function processUnits(metaObj) {
    var softBlue = createOrGetColor("UnitSoftBlue", [230, 234, 246]);
    // 🔍 Récupérer tous les textFrames avec label contenant "unit"
    var allUnits = [];
    for (var i = 0; i < doc.textFrames.length; i++) {
      var tf = doc.textFrames[i];
      if (tf.label.indexOf(UNIT_LABEL) !== -1 && tf.parentPage) {
        allUnits.push({
          frame: tf,
          page: tf.parentPage,
          bounds: tf.geometricBounds
        });
      }
    }
    if (allUnits.length === 0) {
      alert("🚫 Aucun bloc 'unit' trouvé.");
      return;
    }
    // 🔃 Trier tous les blocs par page, puis verticalement
    allUnits.sort(function (a, b) {
      if (a.page.documentOffset !== b.page.documentOffset) {
        return a.page.documentOffset - b.page.documentOffset;
      }
      return a.bounds[0] - b.bounds[0];
    });
    var total = 0;
    var unitsByPage = {};
    for (var i = 0; i < allUnits.length; i++) {
      var pg = allUnits[i].page;
      if (!unitsByPage[pg.name]) unitsByPage[pg.name] = [];
      unitsByPage[pg.name].push(allUnits[i]);
    }
    if (!metaObj.units) metaObj.units = {};
    for (var pageName in unitsByPage) {
      var units = unitsByPage[pageName];
      var page = units[0].page;
      units.sort(function (a, b) { return a.bounds[0] - b.bounds[0]; });
      var pageTop = page.bounds[0];
      var pageLeft = page.bounds[1];
      var pageBottom = page.bounds[2];
      var bottomMargin = page.marginPreferences.bottom;
      var usableBottom = pageBottom - bottomMargin;
      var usableTop = pageTop + TOP_MARGIN_FROM_PAGE;
      var usableHeight = usableBottom - usableTop;
      var marginPrefs = page.marginPreferences;
      var leftMargin = page.bounds[1] + marginPrefs.left;
      var rightMargin = page.bounds[3] - marginPrefs.right;
      // Repositionner chaque bloc "unit" horizontalement aux marges
      for (var i = 0; i < units.length; i++) {
        var unit = units[i];
        var gb = unit.frame.geometricBounds;
        unit.frame.geometricBounds = [gb[0], leftMargin, gb[2], rightMargin];
        unit.bounds = unit.frame.geometricBounds;
      }
      var unitHeights = [];
      var totalUnitHeight = 0;
      for (var i = 0; i < units.length; i++) {
        var h = units[i].bounds[2] - units[i].bounds[0];
        unitHeights.push(h);
        totalUnitHeight += h;
      }
      var totalSpacing = UNIT_TOP_SPACING * (units.length - 1);
      var redBoxHeight = (usableHeight - totalUnitHeight - totalSpacing) / units.length;
      var currentY = usableTop;
      for (var i = 0; i < units.length; i++) {
        var unit = units[i];
        var unitHeight = unitHeights[i];
        var left = unit.bounds[1];
        var right = unit.bounds[3];
        if (i > 0) currentY += UNIT_TOP_SPACING;
        var newBounds = [currentY, left, currentY + unitHeight, right];
        unit.frame.geometricBounds = newBounds;
        // Assign unique label to unit
        var unitId = 'unit_' + generateRandomString(8);
        unit.frame.label = unitId;
        // Clone nmb
        var nbmId = null;
        if (!nmbPrototype) {
          for (var np = 0; np < doc.allPageItems.length; np++) {
            if (doc.allPageItems[np].label === "nmb") {
              nmbPrototype = doc.allPageItems[np];
              break;
            }
          }
        }
        var nmb = null;
        if (nmbPrototype) {
          nmb = nmbPrototype.duplicate(page);
          nmb.bringToFront();
          var unitBounds = newBounds;
          var nmbBounds = nmb.geometricBounds;
          var nmbHeight = nmbBounds[2] - nmbBounds[0];
          var nmbWidth = nmbBounds[3] - nmbBounds[1];
          var nmbTop = unitBounds[0] - 0.2;
          var nmbLeft = unitBounds[1];
          var nmbBottom = nmbTop + nmbHeight;
          var nmbRight = nmbLeft + nmbWidth;
          nmb.geometricBounds = [nmbTop, nmbLeft, nmbBottom, nmbRight];
          var groupItems = nmb.allPageItems;
          for (var g = 0; g < groupItems.length; g++) {
            if (groupItems[g] instanceof TextFrame) {
              if (g === 0) {
                groupItems[g].contents = String(total + 1);
              }
            }
          }
          nbmId = 'nbm_' + metaObj.formation_id + '_' + (total + 1);
          nmb.label = nbmId;
        }
        // Position blue rectangle
        var redTop = newBounds[2] + SPACING_BELOW_UNIT;
        var redBottom = redTop + redBoxHeight;
        var rect = page.rectangles.add();
        rect.geometricBounds = [redTop, left, redBottom, right];
        rect.fillColor = softBlue;
        rect.strokeColor = softBlue;
        var cornerRadius = 0.5;
        rect.topLeftCornerOption = CornerOptions.ROUNDED_CORNER;
        rect.topRightCornerOption = CornerOptions.ROUNDED_CORNER;
        rect.bottomLeftCornerOption = CornerOptions.ROUNDED_CORNER;
        rect.bottomRightCornerOption = CornerOptions.ROUNDED_CORNER;
        rect.topLeftCornerRadius = cornerRadius;
        rect.topRightCornerRadius = cornerRadius;
        rect.bottomLeftCornerRadius = cornerRadius;
        rect.bottomRightCornerRadius = cornerRadius;
        // Assign unique label to rectangle
        var rectId = 'rect_' + metaObj.formation_id + '_' + (total + 1);
        rect.label = rectId;
        // Clone corner
        var cornerId = null;
        if (!cornerPrototype) {
          for (var p = 0; p < doc.allPageItems.length; p++) {
            if (doc.allPageItems[p].label === "corner") {
              cornerPrototype = doc.allPageItems[p];
              break;
            }
          }
        }
        var corner = null;
        if (cornerPrototype) {
          corner = cornerPrototype.duplicate(page);
          var rectBounds = rect.geometricBounds;
          var cornerBounds = corner.geometricBounds;
          var cornerHeight = cornerBounds[2] - cornerBounds[0];
          var cornerWidth = cornerBounds[3] - cornerBounds[1];
          var newTop = rectBounds[2] - cornerHeight;
          var newLeft = rectBounds[3] - cornerWidth;
          var newBottom = rectBounds[2];
          var newRight = rectBounds[3];
          corner.geometricBounds = [newTop, newLeft, newBottom, newRight];
          corner.bringToFront();
          cornerId = 'corner_' + metaObj.formation_id + '_' + (total + 1);
          corner.label = cornerId;
        }
        // Add styled textbox
        var field = page.textBoxes.add({
          geometricBounds: [
            redTop + FIELD_PADDING,
            left + FIELD_PADDING,
            redBottom - FIELD_PADDING,
            right - FIELD_PADDING
          ],
          name: metaObj.formation_id + "_" + (total + 1)
        });
        // Assign unique label to textbox
        var txtbId = 'txtb_' + metaObj.formation_id + '_' + (total + 1);
        field.label = txtbId;
        field.description = "Entrez votre réponse ici...";
        field.appliedFont = "Arial";
        field.fontSize = 10;
        try {
          var tf = field.textFrames[0];
          tf.parentStory.texts[0].appliedFont = app.fonts.item("Arial\tRegular");
          tf.parentStory.texts[0].pointSize = 10;
          tf.contents = " ";
        } catch (e) {}
        // Save unit info in metadata
        metaObj.units[unitId] = {
          index: total + 1,
          page: page.documentOffset + 1,
          position: {
            textframe: newBounds.slice(0),
            rectangle: rect.geometricBounds.slice(0),
            textbox: field.geometricBounds.slice(0)
          },
          rectangleId: rectId,
          texBoxId: txtbId,
          nbmId: nbmId,
          cornerId: cornerId
        };
        total++;
        currentY = redBottom + SPACING_BELOW_UNIT;
      }
    }
    writeMetadata(metaObj);
    alert("✅ " + total + " blocs 'unit' traités avec succès !");
  }

  // === Main logic ===
  metaFrame = findMetaFrame();
  if (metaFrame) {
    metadata = readMetadata();
    SESSION_ID = metadata.formation_id;
    // Prompt user for action
    var action = prompt(
      "Le document a déjà été traité. Que souhaitez-vous faire ?\n\n1: Nettoyer le document\n2: Modifier l'identifiant de formation\n3: Re-générer les blocs\n\nEntrez 1, 2 ou 3 :",
      "1"
    );
    if (action === null) return;
    if (action === "1") {
      cleanUnits(metadata);
      alert("Document nettoyé.");
      return;
    } else if (action === "2") {
      var editedInput = prompt(
        "L’identifiant de formation actuel est : " + SESSION_ID + "\n\nVous pouvez le modifier si désiré :",
        SESSION_ID
      );
      var newValue = String(editedInput).replace(/^\s+|\s+$/g, "");
      if (!newValue) {
        alert("⚠️ Opération annulée.");
        return;
      }
      updateFormationId(metadata, newValue);
      alert("Identifiant de formation mis à jour.");
      return;
    } else if (action === "3") {
      cleanUnits(metadata);
      delete metadata.units;
      writeMetadata(metadata);
      processUnits(metadata);
      return;
    } else {
      alert("Action inconnue ou annulée.");
      return;
    }
  } else {
    // First run: prompt for formation_id
    var userInput = prompt("Veuillez entrer un identifiant de formation :", "");
    var trimmedInput = String(userInput).replace(/^\s+|\s+$/g, "");
    if (!trimmedInput) {
      alert("⚠️ Opération annulée : aucun identifiant saisi.");
      return;
    }
    SESSION_ID = trimmedInput;
    metadata.formation_id = SESSION_ID;
    var page1 = doc.pages[0];
    var pasteboardBounds = page1.bounds;
    var tf = doc.textFrames.add(metadataLayer);
    tf.contents = stringifyJSON(metadata);
    tf.label = SESSION_ID_LABEL;
    var width = 40; // Increased width
    var height = 100; // Increased height
    tf.geometricBounds = [
      pasteboardBounds[0],
      pasteboardBounds[1] - width - 10,
      pasteboardBounds[0] + height,
      pasteboardBounds[1] - 10
    ];
    // Enable hyphenation and allow text to auto-size vertically
    try {
      tf.texts[0].hyphenation = true;
    } catch (e) {}
    tf.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    metaFrame = tf;
    processUnits(metadata);
  }
})();
