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
  var SPACING_BELOW_UNIT = 0; // change to 5 or -2 if you want overlap later

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

  for (var i = 0; i < doc.textFrames.length; i++) {
    var tf = doc.textFrames[i];
    if (tf.label === SESSION_ID_LABEL) {
      metaFrame = tf;
      metadata = parseJSON(tf.contents);
      SESSION_ID = metadata.formation_id;
      break;
    }
  }

  if (SESSION_ID) {
    var editedInput = prompt(
      "L’identifiant de formation actuel est : " + SESSION_ID + "\n\nVous pouvez le modifier si désiré :",
      SESSION_ID
    );
    var newValue = String(editedInput).replace(/^\s+|\s+$/g, "");
    if (!newValue) {
      alert("⚠️ Opération annulée.");
      exit();
    }
    SESSION_ID = newValue;
    metadata.formation_id = SESSION_ID;
    metaFrame.contents = stringifyJSON(metadata);
  } else {
    var userInput = prompt("Veuillez entrer un identifiant de formation :", "");
    var trimmedInput = String(userInput).replace(/^\s+|\s+$/g, "");
    if (!trimmedInput) {
      alert("⚠️ Opération annulée : aucun identifiant saisi.");
      exit();
    }
    SESSION_ID = trimmedInput;
    metadata.formation_id = SESSION_ID;

    var page1 = doc.pages[0];
    var pasteboardBounds = page1.bounds;

    var tf = doc.textFrames.add(metadataLayer);
    tf.contents = stringifyJSON(metadata);
    tf.label = SESSION_ID_LABEL;

    var width = 20;
    var height = 20;

    tf.geometricBounds = [
      pasteboardBounds[0],
      pasteboardBounds[1] - width - 10,
      pasteboardBounds[0] + height,
      pasteboardBounds[1] - 10
    ];
  }


  // === 🧠 Mini JSON helper (for ExtendScript) ===
  function stringifyJSON(obj) {
    var str = '{';
    var first = true;
    for (var key in obj) {
      if (obj.hasOwnProperty(key)) {
        if (!first) str += ',';
        str += '"' + key + '":"' + obj[key].toString().replace(/"/g, '\\"') + '"';
        first = false;
      }
    }
    str += '}';
    return str;
  }

  function parseJSON(str) {
    try {
      // ✋ Use only if you trust the content (safe for internal use)
      return eval('(' + str + ')');
    } catch (e) {
      alert("⚠️ Erreur de lecture des métadonnées JSON.");
      return {};
    }
  }


  // Generate a random string
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
      c.name; // Throws if not found
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

  var softBlue = createOrGetColor("UnitSoftBlue", [230, 234, 246]);

  // 🔍 Récupérer tous les textFrames avec label "unit"
  var allUnits = [];
  for (var i = 0; i < doc.textFrames.length; i++) {
    var tf = doc.textFrames[i];
    if (tf.label === UNIT_LABEL && tf.parentPage) {
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
    return a.bounds[0] - b.bounds[0]; // tri vertical
  });

  var total = 0;

// === Group by page ===
var unitsByPage = {};

for (var i = 0; i < allUnits.length; i++) {
  var pg = allUnits[i].page;
  if (!unitsByPage[pg.name]) unitsByPage[pg.name] = [];
  unitsByPage[pg.name].push(allUnits[i]);
}

// === Process page by page ===
for (var pageName in unitsByPage) {
  var units = unitsByPage[pageName];
  var page = units[0].page;

  // Sort units vertically
  units.sort(function (a, b) {
    return a.bounds[0] - b.bounds[0];
  });

  // Page bounds
  var pageTop = page.bounds[0]; // top Y of this page
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
    // Mettre à jour les bounds pour le calcul plus bas
    unit.bounds = unit.frame.geometricBounds;
  }


  // Calculate total height of all text boxes
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

    if (i > 0) {
      currentY += UNIT_TOP_SPACING;
    }

    // Reposition the text frame to currentY
    var newBounds = [currentY, left, currentY + unitHeight, right];
    unit.frame.geometricBounds = newBounds;

// 📌 Find the first object with label "nmb" once
if (!nmbPrototype) {
  for (var np = 0; np < doc.allPageItems.length; np++) {
    if (doc.allPageItems[np].label === "nmb") {
      nmbPrototype = doc.allPageItems[np];
      break;
    }
  }
}

if (nmbPrototype) {
  var nmb = nmbPrototype.duplicate(page);
  nmb.bringToFront(); // Make sure it's visible

  // Position at top-left of the unit frame
  var unitBounds = newBounds;
  var nmbBounds = nmb.geometricBounds;
  var nmbHeight = nmbBounds[2] - nmbBounds[0];
  var nmbWidth = nmbBounds[3] - nmbBounds[1];

  var nmbTop = unitBounds[0] - 0.2;
  var nmbLeft = unitBounds[1];
  var nmbBottom = nmbTop + nmbHeight;
  var nmbRight = nmbLeft + nmbWidth;

  nmb.geometricBounds = [nmbTop, nmbLeft, nmbBottom, nmbRight];

  // 🧠 Update text inside the group
  var groupItems = nmb.allPageItems;
  for (var g = 0; g < groupItems.length; g++) {
    if (groupItems[g] instanceof TextFrame) {
      if (g === 0) {
        groupItems[g].contents = String(i + 1);
      }

    }
  }
}

    // Position red rectangle right after
    var redTop = newBounds[2] + SPACING_BELOW_UNIT;
    var redBottom = redTop + redBoxHeight;

    var rect = page.rectangles.add();
    rect.geometricBounds = [redTop, left, redBottom, right];
    rect.fillColor = softBlue;
    rect.strokeColor = softBlue;

    var cornerRadius = 0.5; // in points

    rect.topLeftCornerOption = CornerOptions.ROUNDED_CORNER;
    rect.topRightCornerOption = CornerOptions.ROUNDED_CORNER;
    rect.bottomLeftCornerOption = CornerOptions.ROUNDED_CORNER;
    rect.bottomRightCornerOption = CornerOptions.ROUNDED_CORNER;

    rect.topLeftCornerRadius = cornerRadius;
    rect.topRightCornerRadius = cornerRadius;
    rect.bottomLeftCornerRadius = cornerRadius;
    rect.bottomRightCornerRadius = cornerRadius;

    // 📌 Find the first object with label "corner" only once
    if (!cornerPrototype) {
      for (var p = 0; p < doc.allPageItems.length; p++) {
        if (doc.allPageItems[p].label === "corner") {
          cornerPrototype = doc.allPageItems[p];
          break;
        }
      }
    }

  if (cornerPrototype) {
    var corner = cornerPrototype.duplicate(page);

    // Align it to the bottom right of the rect
    var rectBounds = rect.geometricBounds; // [top, left, bottom, right]
    var cornerBounds = corner.geometricBounds;
    var cornerHeight = cornerBounds[2] - cornerBounds[0];
    var cornerWidth = cornerBounds[3] - cornerBounds[1];

    var newTop = rectBounds[2] - cornerHeight;
    var newLeft = rectBounds[3] - cornerWidth;
    var newBottom = rectBounds[2];
    var newRight = rectBounds[3];

    corner.geometricBounds = [newTop, newLeft, newBottom, newRight];
    corner.bringToFront(); // ✅ Assure que l'objet "corner" soit au-dessus de tout

  }

    var field = page.textBoxes.add({
      geometricBounds: [
        redTop + FIELD_PADDING,
        left + FIELD_PADDING,
        redBottom - FIELD_PADDING,
        right - FIELD_PADDING
      ],
      name: SESSION_ID + "_" + (total + 1)
    });

    // Visual properties
    field.description = "Entrez votre réponse ici..."
    field.appliedFont = "Arial";
    field.fontSize = 10;

    try {
      var tf = field.textFrames[0];
      tf.parentStory.texts[0].appliedFont = app.fonts.item("Arial\tRegular");
      tf.parentStory.texts[0].pointSize = 10;
      tf.contents = " ";
    } catch (e) {
      $.writeln("⚠️ Font issue: " + e);
    }

    total++;
    currentY = redBottom + SPACING_BELOW_UNIT;
  }
}
  alert("✅ " + total + " blocs 'unit' traités avec succès !");
})();
