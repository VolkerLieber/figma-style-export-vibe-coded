// ─── Helpers ────────────────────────────────────────────────────────────────

function rgbToHex(r, g, b) {
  var toHex = function(v) { return Math.round(v * 255).toString(16).padStart(2, '0'); };
  return '#' + toHex(r) + toHex(g) + toHex(b);
}

function rgbaToHex(r, g, b, a) {
  var toHex = function(v) { return Math.round(v * 255).toString(16).padStart(2, '0'); };
  var alpha = (a !== undefined && a < 1) ? toHex(a) : '';
  return '#' + toHex(r) + toHex(g) + toHex(b) + alpha;
}

function parseName(name) {
  var parts = name.split('/');
  return {
    group: parts.length > 1 ? parts.slice(0, -1).join('/') : null,
    token: parts[parts.length - 1],
    path:  parts,
  };
}

function shallowMerge(target) {
  for (var i = 1; i < arguments.length; i++) {
    var src = arguments[i];
    if (!src) continue;
    var keys = Object.keys(src);
    for (var k = 0; k < keys.length; k++) {
      target[keys[k]] = src[keys[k]];
    }
  }
  return target;
}

// Wrap a promise with a timeout so a stalled Figma IPC call never hangs export forever.
function withTimeout(promise, ms) {
  return new Promise(function(resolve) {
    var timer = setTimeout(function() { resolve(null); }, ms);
    promise.then(function(v) { clearTimeout(timer); resolve(v); },
                 function()  { clearTimeout(timer); resolve(null); });
  });
}

async function resolveVariableAlias(alias) {
  if (!alias || alias.type !== 'VARIABLE_ALIAS') return null;
  try {
    var variable = await withTimeout(figma.variables.getVariableByIdAsync(alias.id), 500);
    return variable ? { name: variable.name, id: variable.id } : null;
  } catch (e) {
    return null;
  }
}

// Factory: captures bv at call-time so the async body never sees a stale reference
// when the outer for-loop has advanced to the next iteration.
function makeResolveField(bv) {
  return async function(field) {
    var aliases = bv[field];
    if (!aliases) return null;
    var alias = Array.isArray(aliases) ? aliases[0] : aliases;
    return resolveVariableAlias(alias);
  };
}

// ─── Text Styles ────────────────────────────────────────────────────────────

async function extractTextStyles() {
  var styles = await figma.getLocalTextStylesAsync();
  var result = [];

  for (var i = 0; i < styles.length; i++) {
    var style = styles[i];
    var bv = style.boundVariables || {};

    // Post progress every 10 styles so the UI doesn't appear frozen on large files
    if (i % 10 === 0 || i === styles.length - 1) {
      figma.ui.postMessage({ type: 'STATUS', message: 'Text styles ' + (i + 1) + '/' + styles.length + '…' });
    }

    var resolveField = makeResolveField(bv);

    var nameMeta = parseName(style.name);

    result.push(shallowMerge({
      id:               style.key,
      name:             style.name,
      description:      style.description || '',
      fontSize:         style.fontSize,
      fontFamily:       style.fontName ? style.fontName.family : null,
      fontStyle:        style.fontName ? style.fontName.style  : null,
      textDecoration:   style.textDecoration,
      letterSpacing:    style.letterSpacing,
      lineHeight:       style.lineHeight,
      leadingTrim:      style.leadingTrim,
      paragraphIndent:  style.paragraphIndent,
      paragraphSpacing: style.paragraphSpacing,
      listSpacing:      style.listSpacing,
      hangingPunctuation: style.hangingPunctuation,
      hangingList:      style.hangingList,
      textCase:         style.textCase,
      boundVariables: {
        fontSize:         await resolveField('fontSize'),
        fontFamily:       await resolveField('fontFamily'),
        fontStyle:        await resolveField('fontStyle'),
        letterSpacing:    await resolveField('letterSpacing'),
        lineHeight:       await resolveField('lineHeight'),
        paragraphIndent:  await resolveField('paragraphIndent'),
        paragraphSpacing: await resolveField('paragraphSpacing'),
      },
    }, nameMeta));
  }

  return result;
}

// ─── Paint Styles ───────────────────────────────────────────────────────────

async function extractPaintStyles() {
  var styles = await figma.getLocalPaintStylesAsync();
  var result = [];

  for (var i = 0; i < styles.length; i++) {
    var style = styles[i];
    var paints = [];

    for (var j = 0; j < style.paints.length; j++) {
      var paint = style.paints[j];
      var base = {
        type:    paint.type,
        visible: paint.visible !== false,
        opacity: paint.opacity !== undefined ? paint.opacity : 1,
      };

      if (paint.type === 'SOLID') {
        var hex = rgbaToHex(paint.color.r, paint.color.g, paint.color.b, paint.opacity !== undefined ? paint.opacity : 1);
        var boundColorVar = (paint.boundVariables && paint.boundVariables.color)
          ? await resolveVariableAlias(paint.boundVariables.color)
          : null;
        paints.push(shallowMerge({}, base, { hex: hex, boundVariable: boundColorVar }));

      } else if (paint.type.indexOf('GRADIENT') === 0) {
        var stops = [];
        var gradientStops = paint.gradientStops || [];
        for (var s = 0; s < gradientStops.length; s++) {
          var stop = gradientStops[s];
          var stopHex = rgbToHex(stop.color.r, stop.color.g, stop.color.b);
          var stopVar = (stop.boundVariables && stop.boundVariables.color)
            ? await resolveVariableAlias(stop.boundVariables.color)
            : null;
          stops.push({ position: stop.position, hex: stopHex, boundVariable: stopVar });
        }
        paints.push(shallowMerge({}, base, { gradientStops: stops }));

      } else {
        paints.push(base);
      }
    }

    result.push(shallowMerge({
      id:          style.key,
      name:        style.name,
      description: style.description || '',
      paints:      paints,
    }, parseName(style.name)));
  }

  return result;
}

// ─── Effect Styles ──────────────────────────────────────────────────────────

async function extractEffectStyles() {
  var styles = await figma.getLocalEffectStylesAsync();
  var result = [];

  for (var i = 0; i < styles.length; i++) {
    var style = styles[i];
    var effects = [];

    for (var j = 0; j < style.effects.length; j++) {
      var effect = style.effects[j];
      var e = { type: effect.type, visible: effect.visible !== false };
      if (effect.color)                  e.color     = rgbaToHex(effect.color.r, effect.color.g, effect.color.b, effect.color.a);
      if (effect.offset)                 e.offset    = effect.offset;
      if (effect.radius !== undefined)   e.radius    = effect.radius;
      if (effect.spread !== undefined)   e.spread    = effect.spread;
      if (effect.blendMode)              e.blendMode = effect.blendMode;
      effects.push(e);
    }

    result.push(shallowMerge({
      id:          style.key,
      name:        style.name,
      description: style.description || '',
      effects:     effects,
    }, parseName(style.name)));
  }

  return result;
}

// ─── Variables ──────────────────────────────────────────────────────────────

async function extractVariables() {
  var collections = await figma.variables.getLocalVariableCollectionsAsync();
  var allVars     = await figma.variables.getLocalVariablesAsync();

  var varById = {};
  for (var i = 0; i < allVars.length; i++) {
    varById[allVars[i].id] = allVars[i];
  }

  // Cache for external variable lookups to avoid redundant getVariableByIdAsync
  // calls — a file with 50 Token Sizes vars × 3 modes hits the same handful
  // of external IDs repeatedly without this.
  var extVarCache = {};

  // resolveValue is async: for VARIABLE_ALIAS values pointing to external
  // (library) variables not in varById, we attempt getVariableByIdAsync to
  // retrieve the variable name — critical for Token Sizes aliases like
  // Desktop→size/text-6xl which live in an external file.
  async function resolveValue(value, resolvedType) {
    if (value && typeof value === 'object' && value.type === 'VARIABLE_ALIAS') {
      var ref = varById[value.id];
      if (ref) return { $alias: ref.name, $aliasId: ref.id };
      // External variable — check cache first, then API
      if (extVarCache[value.id] !== undefined) {
        return extVarCache[value.id];
      }
      try {
        var extVar = await withTimeout(figma.variables.getVariableByIdAsync(value.id), 500);
        var resolved = extVar ? { $alias: extVar.name, $aliasId: extVar.id } : { $alias: value.id };
        extVarCache[value.id] = resolved;
        return resolved;
      } catch (e) {
        extVarCache[value.id] = { $alias: value.id };
        return { $alias: value.id };
      }
    }
    if (resolvedType === 'COLOR' && value && typeof value === 'object') {
      return rgbaToHex(value.r, value.g, value.b, value.a !== undefined ? value.a : 1);
    }
    return value;
  }

  var result = [];

  for (var c = 0; c < collections.length; c++) {
    var collection = collections[c];
    var variables  = [];

    figma.ui.postMessage({ type: 'STATUS', message: 'Variables: ' + collection.name + '…' });

    for (var v = 0; v < collection.variableIds.length; v++) {
      var variable = varById[collection.variableIds[v]];
      if (!variable) continue;

      var valuesByMode = {};
      var modeIds = Object.keys(variable.valuesByMode);
      for (var m = 0; m < modeIds.length; m++) {
        var modeId   = modeIds[m];
        var modeName = modeId;
        for (var mm = 0; mm < collection.modes.length; mm++) {
          if (collection.modes[mm].modeId === modeId) { modeName = collection.modes[mm].name; break; }
        }
        valuesByMode[modeName] = await resolveValue(variable.valuesByMode[modeId], variable.resolvedType);
      }

      variables.push(shallowMerge({
        id:           variable.id,
        name:         variable.name,
        resolvedType: variable.resolvedType,
        description:  variable.description || '',
        scopes:       variable.scopes,
        codeSyntax:   variable.codeSyntax,
        valuesByMode: valuesByMode,
      }, parseName(variable.name)));
    }

    result.push({
      id:            collection.id,
      name:          collection.name,
      defaultModeId: collection.defaultModeId,
      modes:         collection.modes,
      variables:     variables,
    });
  }

  return result;
}

// ─── Library Variables (descriptors only, no import) ────────────────────────

async function extractLibraryVariables() {
  if (!figma.teamLibrary) return [];

  var libCollections;
  try {
    libCollections = await figma.teamLibrary.getAvailableLibraryVariableCollectionsAsync();
  } catch (e) {
    return [];
  }

  if (!libCollections || !libCollections.length) return [];

  var result = [];

  for (var i = 0; i < libCollections.length; i++) {
    var libCollection = libCollections[i];
    var libVars;
    try {
      libVars = await figma.teamLibrary.getVariablesInLibraryCollectionAsync(libCollection.key);
    } catch (e) {
      continue;
    }

    var variables = [];
    for (var j = 0; j < libVars.length; j++) {
      var v = libVars[j];
      variables.push(shallowMerge({
        key:          v.key,
        name:         v.name,
        resolvedType: v.resolvedType,
      }, parseName(v.name)));
    }

    result.push({
      id:          libCollection.key,
      name:        libCollection.name,
      libraryName: libCollection.libraryName,
      source:      'library',
      valuesResolved: false,
      variables:   variables,
    });
  }

  return result;
}

// ─── Library Variable Values (import → read, no destructive cleanup) ─────────
//
// Safety model:
//   - importVariableByKeyAsync is idempotent: calling it on an already-present
//     variable returns the existing one without duplicating it.
//   - We never call variable.remove(). Remote variables Figma introduced have
//     remote:true and remove() is only documented for local variables. Figma
//     garbage-collects unused remote references automatically.
//   - We track which variables were pre-existing vs newly introduced so the
//     caller knows what changed, but we do not touch pre-existing ones at all.
//   - A small throttle between imports avoids 429 rate-limit errors on large libs.

async function resolveLibraryVariableValues(libraryVariables) {
  // Build a set of variable keys already present in the document before we start.
  // importVariableByKeyAsync returns the existing object for these — no new write.
  var existingVars = await figma.variables.getLocalVariablesAsync();
  var existingKeys = {};
  for (var e = 0; e < existingVars.length; e++) {
    existingKeys[existingVars[e].key] = true;
  }

  var result = [];

  for (var c = 0; c < libraryVariables.length; c++) {
    var col = libraryVariables[c];
    var resolvedVars  = [];
    var newlyImported = 0; // informational only, no removal

    figma.ui.postMessage({
      type: 'STATUS',
      message: 'Resolving "' + col.name + '" (' + (c + 1) + '/' + libraryVariables.length + ')\u2026'
    });

    for (var v = 0; v < col.variables.length; v++) {
      var descriptor  = col.variables[v];
      var wasExisting = existingKeys[descriptor.key] === true;
      var imported;

      try {
        imported = await figma.variables.importVariableByKeyAsync(descriptor.key);
      } catch (e) {
        // Not published or import failed — keep descriptor without values
        resolvedVars.push(descriptor);
        // Small delay even on failure to avoid hammering the API
        await new Promise(function(r) { setTimeout(r, 50); });
        continue;
      }

      if (!wasExisting) newlyImported++;

      // Read valuesByMode. Note: this will not resolve cross-library aliases
      // automatically — alias IDs are preserved as-is for transparency.
      // Build a modeId → modeName map from the imported variable's collection.
      // Note: getVariableCollectionByIdAsync only returns LOCAL collections.
      // For truly external library variables their collection is remote and will
      // not be found — modeIdToName stays empty and keys fall back to raw modeId strings.
      var importedCollection = null;
      try {
        importedCollection = await withTimeout(figma.variables.getVariableCollectionByIdAsync(imported.variableCollectionId), 500);
      } catch (e) { /* remote collection — not locally accessible, fall back to modeId keys */ }
      var modeIdToName = {};
      if (importedCollection) {
        for (var mc = 0; mc < importedCollection.modes.length; mc++) {
          modeIdToName[importedCollection.modes[mc].modeId] = importedCollection.modes[mc].name;
        }
      }

      var valuesByMode = {};
      var modeIds = Object.keys(imported.valuesByMode);
      for (var m = 0; m < modeIds.length; m++) {
        var modeId   = modeIds[m];
        var modeName = modeIdToName[modeId] || modeId; // fall back to ID if name unavailable
        var raw      = imported.valuesByMode[modeId];
        var val;
        if (raw && typeof raw === 'object' && raw.type === 'VARIABLE_ALIAS') {
          // Preserve alias reference — don't silently resolve to a value
          val = { $alias: raw.id };
        } else if (imported.resolvedType === 'COLOR' && raw && typeof raw === 'object') {
          val = rgbaToHex(raw.r, raw.g, raw.b, raw.a !== undefined ? raw.a : 1);
        } else {
          val = raw;
        }
        valuesByMode[modeName] = val;
      }

      resolvedVars.push(shallowMerge({}, descriptor, {
        valuesByMode: valuesByMode,
        description:  imported.description || '',
        codeSyntax:   imported.codeSyntax,
        wasExisting:  wasExisting,
      }));

      // Throttle: 80ms between imports to avoid rate limiting on large collections
      await new Promise(function(r) { setTimeout(r, 80); });
    }

    result.push(shallowMerge({}, col, {
      valuesResolved: true,
      importStats: { newlyImported: newlyImported },
      variables: resolvedVars,
    }));
  }

  return result;
}

// ─── Main ───────────────────────────────────────────────────────────────────

figma.showUI(__html__, { width: 480, height: 620, title: 'Token Exporter' });

figma.ui.onmessage = async function(msg) {

  // ── Initial export ──────────────────────────────────────────────────────────
  if (msg.type === 'EXPORT') {
    try {
      figma.ui.postMessage({ type: 'STATUS', message: 'Collecting text styles\u2026' });
      var textStyles = await extractTextStyles();

      figma.ui.postMessage({ type: 'STATUS', message: 'Collecting paint styles\u2026' });
      var paintStyles = await extractPaintStyles();

      figma.ui.postMessage({ type: 'STATUS', message: 'Collecting effect styles\u2026' });
      var effectStyles = await extractEffectStyles();

      figma.ui.postMessage({ type: 'STATUS', message: 'Collecting variables\u2026' });
      var variables = await extractVariables();

      figma.ui.postMessage({ type: 'STATUS', message: 'Collecting library variables\u2026' });
      var libraryVariables = await extractLibraryVariables();

      var payload = {
        meta: {
          exportedAt:   new Date().toISOString(),
          figmaFileKey: figma.fileKey || null,
          counts: {
            textStyles:          textStyles.length,
            paintStyles:         paintStyles.length,
            effectStyles:        effectStyles.length,
            variableCollections: variables.length,
            variables:           variables.reduce(function(acc, c) { return acc + c.variables.length; }, 0),
            libraryCollections:  libraryVariables.length,
            libraryVariables:    libraryVariables.reduce(function(acc, c) { return acc + c.variables.length; }, 0),
          },
        },
        textStyles:       textStyles,
        paintStyles:      paintStyles,
        effectStyles:     effectStyles,
        variables:        variables,
        libraryVariables: libraryVariables,
      };

      figma.ui.postMessage({ type: 'RESULT', payload: payload });
    } catch (err) {
      figma.ui.postMessage({ type: 'ERROR', message: err.message || String(err) });
    }
  } else if (msg.type === 'RESOLVE_LIB_VARS') { // ── Resolve library variable values (opt-in, user confirmed) ──
    try {
      // Validate incoming data before passing to the resolver
      if (!Array.isArray(msg.libraryVariables)) {
        figma.ui.postMessage({ type: 'ERROR', message: 'Invalid library variable data received.' });
        return;
      }
      var validCollections = msg.libraryVariables.filter(function(col) {
        return col && Array.isArray(col.variables);
      });
      var resolved = await resolveLibraryVariableValues(validCollections);
      figma.ui.postMessage({ type: 'LIB_VARS_RESOLVED', libraryVariables: resolved });
    } catch (err) {
      figma.ui.postMessage({ type: 'ERROR', message: err.message || String(err) });
    }
  }

};
