// ★対象のスプレッドシートIDをここに設定してください
const SPREADSHEET_ID = '1npq_dPbsnNWGKgRgO3BOF-JB1QfHwQ_EHhJPlfFtsZs';

/**
 * スプレッドシートを取得するヘルパー関数
 */
function getTargetSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== '1npq_dPbsnNWGKgRgO3BOF-JB1QfHwQ_EHhJPlfFtsZs') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 外部（GitHub Pages等）からの fetch 通信を受け取るAPIエンドポイント
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : 'getTopBannerMap';
  const charName = (e && e.parameter && e.parameter.charName) ? String(e.parameter.charName) : '';

  let result = {};

  try {
    if (action === 'getTopBannerMap') {
      result = getTopBannerMap();
    } else if (action === 'getUpdateHistory') { // ★ 追加
      result = getUpdateHistory();
    } else if (action === 'getCharacterDataGroupedByUnit') {
      result = getCharacterDataGroupedByUnit();
    } else if (action === 'getHolomenTableData') {
      result = getHolomenTableData();
    } else if (action === 'getCollectionPageData') {
      result = getCollectionPageData();
    } else if (action === 'getCharacterDetailsByName') {
      result = getCharacterDetailsByName(charName);
    }
  } catch (err) {
    result = { error: true, message: err.message };
  }

  const output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * 安全にシートの値を取得
 */
function getSafeSheetValues(sheetName) {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    return sheet.getDataRange().getValues();
  } catch (e) {
    console.warn('シート読み込みスキップ: ' + sheetName, e);
    return [];
  }
}

/**
 * シートのデータをオブジェクトの配列形式で安全に取得するヘルパー関数
 */
function getSheetDataAsObjects(sheetName) {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0].map(h => String(h).trim());
    const rows = [];
    
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      let hasValue = false;
      for (let j = 0; j < headers.length; j++) {
        const val = data[i][j];
        obj[headers[j]] = val;
        if (val !== '' && val !== null && val !== undefined) hasValue = true;
      }
      if (hasValue) rows.push(obj);
    }
    return rows;
  } catch (e) {
    console.warn('マスターシート読み込みスキップ: ' + sheetName, e);
    return [];
  }
}

/**
 * Topページ用：bannerシートから画像URLマップを取得
 */
function getTopBannerMap() {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName('banner');
    if (!sheet) return {};

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return {};

    const headers = values[0].map(h => String(h || '').trim());
    const nameIdx = headers.findIndex(h => ['name', '機能名', 'title', '項目'].includes(h));
    const urlIdx = headers.findIndex(h => ['banner_url', 'bannerurl', 'url', 'image'].includes(h));

    const bannerMap = {};
    for (let i = 1; i < values.length; i++) {
      const name = nameIdx !== -1 ? String(values[i][nameIdx] || '').trim() : String(values[i][0] || '').trim();
      const url = urlIdx !== -1 ? String(values[i][urlIdx] || '').trim() : String(values[i][1] || '').trim();
      if (name) bannerMap[name] = url;
    }
    return bannerMap;
  } catch (err) {
    console.error('バナー画像取得エラー:', err);
    return {};
  }
}

/**
 * タレント画面用：ユニットごとのキャラクターデータ取得
 */
function getCharacterDataGroupedByUnit() {
  const ss = getTargetSpreadsheet();
  
  const tagValues = getSafeSheetValues('tag');
  const tagUrlMap = {};
  if (tagValues.length >= 2) {
    const tHeaders = tagValues[0].map(h => String(h).trim());
    const tagIdx = tHeaders.findIndex(h => ['tag', 'tag_name'].includes(h));
    const tagUrlIdx = tHeaders.indexOf('tag_url');
    if (tagIdx !== -1 && tagUrlIdx !== -1) {
      for (let i = 1; i < tagValues.length; i++) {
        const tVal = tagValues[i][tagIdx];
        const uVal = tagValues[i][tagUrlIdx];
        if (tVal) tagUrlMap[String(tVal).trim()] = uVal ? String(uVal).trim() : '';
      }
    }
  }

  const unitValues = getSafeSheetValues('unit');
  const unitInfoMap = {};
  if (unitValues.length >= 2) {
    const uHeaders = unitValues[0].map(h => String(h).trim());
    const uNameIdx = uHeaders.findIndex(h => ['unit', 'unit_name'].includes(h));
    const uPriorityIdx = uHeaders.indexOf('priority');
    const uTagNameIdx = uHeaders.indexOf('tag_name');

    for (let i = 1; i < unitValues.length; i++) {
      const uName = unitValues[i][uNameIdx !== -1 ? uNameIdx : 0];
      const uPrio = uHeaders.indexOf('priority') !== -1 ? unitValues[i][uPriorityIdx] : (i * 10);
      const uTagName = uTagNameIdx !== -1 ? unitValues[i][uTagNameIdx] : '';

      if (uName) {
        const trimmedUnitName = String(uName).trim();
        const trimmedTagName = uTagName ? String(uTagName).trim() : '';
        unitInfoMap[trimmedUnitName] = {
          priority: Number(uPrio) || (i * 10),
          tagUrl: tagUrlMap[trimmedTagName] || ''
        };
      }
    }
  }

  const holomenSheet = ss.getSheetByName('ホロメン');
  const charSearchWordsMap = {};

  if (holomenSheet) {
    const hValues = holomenSheet.getDataRange().getValues();
    if (hValues.length >= 2) {
      const hHeaders = hValues[0].map(h => String(h).trim());
      const findHIdx = (names) => hHeaders.findIndex(h => names.includes(h));

      const hNameIdx = findHIdx(['character', 'キャラクター名', 'character_name', '名前']);
      const hAliasIdx = findHIdx(['alias', '二つ名', '異名', 'メンバー名']);
      const hSearchWordIdx = findHIdx(['search_word', 'searchword', 'search_words', 'searchwords']);

      for (let i = 1; i < hValues.length; i++) {
        const row = hValues[i];
        if (row[0] === "" || row[0] === null || row[0] === undefined) continue;

        const cName = hNameIdx !== -1 ? String(row[hNameIdx] || '').trim() : '';
        if (!cName) continue;

        if (!charSearchWordsMap[cName]) charSearchWordsMap[cName] = new Set();

        const aliasVal = hAliasIdx !== -1 ? String(row[hAliasIdx] || '').trim() : '';
        const swVal = hSearchWordIdx !== -1 ? String(row[hSearchWordIdx] || '').trim() : '';

        if (aliasVal) charSearchWordsMap[cName].add(aliasVal);
        if (swVal) {
          swVal.split(/[,、]/).forEach(w => {
            const trimmed = w.trim();
            if (trimmed) charSearchWordsMap[cName].add(trimmed);
          });
        }
      }
    }
  }

  const charSheet = ss.getSheetByName('character');
  if (!charSheet) {
    throw new Error('「character」シートが見つかりませんでした。');
  }

  const charValues = charSheet.getDataRange().getValues();
  if (charValues.length < 2) return [];

  const headers = charValues[0].map(h => String(h).trim());
  const findIdx = (names) => headers.findIndex(h => names.includes(h));

  const unitIdx = findIdx(['unit', 'ユニット']);
  const iconIdx = findIdx(['icon', 'アイコン']);
  const nameIdx = findIdx(['character', 'キャラクター名', 'character_name', '名前']);
  const aliasIdx = findIdx(['alias', '二つ名', '異名']);
  const plateIdx = findIdx(['character_plate', 'characterplate', 'char_plate']);
  const releaseFlagIdx = findIdx(['release_flag', 'releaseflag', 'release']);
  const searchWordIdx = findIdx(['search_word', 'searchword', 'search_words', 'searchwords']);

  const groupedDataMap = {};
  const processedCharPerUnit = {};

  charValues.slice(1).forEach(row => {
    if (row[0] === "" || row[0] === null || row[0] === undefined) return;

    if (releaseFlagIdx !== -1) {
      const flagVal = row[releaseFlagIdx];
      if (flagVal === false || String(flagVal).trim().toLowerCase() === 'false' || flagVal === 0) {
        return;
      }
    }

    const characterName = nameIdx !== -1 ? String(row[nameIdx]).trim() : '';
    if (!characterName) return;

    const iconUrl = iconIdx !== -1 ? row[iconIdx] : '';
    const alias = aliasIdx !== -1 ? row[aliasIdx] : '';
    const characterPlate = plateIdx !== -1 ? row[plateIdx] : '';
    const rawUnitStr = (unitIdx !== -1 && row[unitIdx]) ? String(row[unitIdx]) : 'その他';

    const searchWordSet = charSearchWordsMap[characterName] || new Set();
    const selfSw = searchWordIdx !== -1 ? String(row[searchWordIdx] || '').trim() : '';
    if (selfSw) {
      selfSw.split(/[,、]/).forEach(w => { if (w.trim()) searchWordSet.add(w.trim()); });
    }

    const units = rawUnitStr.split(/[,、]/).map(u => u.trim()).filter(u => u.length > 0);
    const targetUnits = units.length > 0 ? units : ['その他'];

    targetUnits.forEach(unitName => {
      if (!groupedDataMap[unitName]) groupedDataMap[unitName] = [];
      if (!processedCharPerUnit[unitName]) processedCharPerUnit[unitName] = new Set();

      if (processedCharPerUnit[unitName].has(characterName)) return;
      processedCharPerUnit[unitName].add(characterName);

      groupedDataMap[unitName].push({
        id: row[0],
        character: characterName,
        icon: iconUrl,
        alias: alias,
        unit: unitName,
        character_plate: characterPlate,
        search_words: Array.from(searchWordSet)
      });
    });
  });

  const sortedUnits = Object.keys(groupedDataMap).sort((a, b) => {
    const prioA = unitInfoMap[a] ? unitInfoMap[a].priority : 999;
    const prioB = unitInfoMap[b] ? unitInfoMap[b].priority : 999;
    return prioA - prioB;
  });

  return sortedUnits.map(unitName => ({
    unitName: unitName,
    priority: unitInfoMap[unitName] ? unitInfoMap[unitName].priority : 999,
    tagUrl: unitInfoMap[unitName] ? unitInfoMap[unitName].tagUrl : '',
    characters: groupedDataMap[unitName]
  }));
}

/**
 * メンバー画面用データ取得
 */
function getHolomenTableData() {
  try {
    const ss = getTargetSpreadsheet();

    const gValues = getSafeSheetValues('global_icon');
    const globalIcons = {};
    if (gValues.length >= 2) {
      const gHeaders = gValues[0].map(h => String(h).trim());
      const nameIdx = gHeaders.findIndex(h => ['name', 'icon_name', 'type', 'key'].includes(h));
      const urlIdx = gHeaders.indexOf('icon_url');
      for (let i = 1; i < gValues.length; i++) {
        const k = nameIdx !== -1 ? String(gValues[i][nameIdx]).trim() : String(gValues[i][0]).trim();
        const u = urlIdx !== -1 ? gValues[i][urlIdx] : gValues[i][1];
        if (k) globalIcons[k] = u;
      }
    }

    const tValues = getSafeSheetValues('tag');
    const tagUrlMap = {};
    if (tValues.length >= 2) {
      const tHeaders = tValues[0].map(h => String(h).trim());
      const tagIdx = tHeaders.findIndex(h => ['tag', 'tag_name'].includes(h));
      const tagUrlIdx = tHeaders.indexOf('tag_url');
      if (tagIdx !== -1 && tagUrlIdx !== -1) {
        for (let i = 1; i < tValues.length; i++) {
          const tVal = tValues[i][tagIdx];
          const uVal = tValues[i][tagUrlIdx];
          if (tVal) tagUrlMap[String(tVal).trim()] = uVal ? String(uVal).trim() : '';
        }
      }
    }

    const uValues = getSafeSheetValues('unit');
    const unitTagMap = {};
    if (uValues.length >= 2) {
      const uHeaders = uValues[0].map(h => String(h).trim());
      const uNameIdx = uHeaders.findIndex(h => ['unit', 'unit_name'].includes(h));
      const uTagNameIdx = uHeaders.indexOf('tag_name');
      for (let i = 1; i < uValues.length; i++) {
        const uName = uValues[i][uNameIdx !== -1 ? uNameIdx : 0];
        const uTagName = uTagNameIdx !== -1 ? uValues[i][uTagNameIdx] : '';
        if (uName) {
          const tagName = uTagName ? String(uTagName).trim() : '';
          unitTagMap[String(uName).trim()] = tagUrlMap[tagName] || '';
        }
      }
    }

    const charSheet = ss.getSheetByName('ホロメン') || ss.getSheetByName('character');
    if (!charSheet) {
      throw new Error('「ホロメン」または「character」シートが存在しません。');
    }

    const values = charSheet.getDataRange().getValues();
    if (values.length < 2) return { rows: [], globalIcons: globalIcons, tagUrlMap: tagUrlMap, masterData: {} };

    const headers = values[0].map(h => String(h).trim());
    const releaseFlagIdx = headers.findIndex(h => ['release_flag', 'releaseflag', 'release'].includes(h));

    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[0] === "" || row[0] === null || row[0] === undefined) continue;

      if (releaseFlagIdx !== -1) {
        const flagVal = row[releaseFlagIdx];
        if (flagVal === false || String(flagVal).trim().toLowerCase() === 'false' || flagVal === 0) continue;
      }

      const rowObj = {};
      headers.forEach((h, idx) => {
        if (h) rowObj[h] = row[idx];
      });

      const rawUnits = rowObj['unit'] ? String(rowObj['unit']).split(/[,、]/).map(u => u.trim()) : [];
      rowObj['unit_tag_urls'] = rawUnits.map(u => unitTagMap[u]).filter(url => Boolean(url));

      rows.push(rowObj);
    }

    rows.sort((a, b) => {
      const rarityA = Number(a['rarity']) || 0;
      const rarityB = Number(b['rarity']) || 0;
      if (rarityA !== rarityB) return rarityB - rarityA;

      const releaseA = Number(String(a['release_at'] || a['releaseat']).replace(/\D/g, '')) || 0;
      const releaseB = Number(String(b['release_at'] || b['releaseat']).replace(/\D/g, '')) || 0;
      if (releaseA !== releaseB) return releaseB - releaseA;

      const idA = Number(a['id']) || 0;
      const idB = Number(b['id']) || 0;
      return idA - idB;
    });

    const masterData = {
      skill_effect: getSheetDataAsObjects('skill_effect'),
      condition_type: getSheetDataAsObjects('condition_type'),
      skill_conditions: getSheetDataAsObjects('skill_conditions'),
      skill_target_type: getSheetDataAsObjects('skill_target_type'),
      skill_target: getSheetDataAsObjects('skill_target')
    };

    return {
      rows: rows,
      globalIcons: globalIcons,
      tagUrlMap: tagUrlMap,
      masterData: masterData
    };
  } catch (err) {
    throw new Error('サーバー処理エラー: ' + err.message);
  }
}

/**
 * コレクション画面用データ取得
 */
function getCollectionPageData() {
  try {
    const ss = getTargetSpreadsheet();

    const gValues = getSafeSheetValues('global_icon');
    const globalIcons = {};
    if (gValues.length >= 2) {
      const gHeaders = gValues[0].map(h => String(h).trim());
      const nameIdx = gHeaders.findIndex(h => ['name', 'icon_name', 'type', 'key'].includes(h));
      const urlIdx = gHeaders.indexOf('icon_url');
      for (let i = 1; i < gValues.length; i++) {
        const k = nameIdx !== -1 ? String(gValues[i][nameIdx]).trim() : String(gValues[i][0]).trim();
        const u = urlIdx !== -1 ? gValues[i][urlIdx] : gValues[i][1];
        if (k) globalIcons[k] = u;
      }
    }

    const rValues = getSafeSheetValues('rarity');
    const rarityUrlMap = {};
    if (rValues.length >= 2) {
      const rHeaders = rValues[0].map(h => String(h).trim());
      const rarityIdx = rHeaders.findIndex(h => ['rarity', 'レアリティ', 'レア度'].includes(h));
      const urlIdx = rHeaders.findIndex(h => ['rarity_url', 'rarityurl', 'icon_url', 'url'].includes(h));

      for (let i = 1; i < rValues.length; i++) {
        const rVal = rarityIdx !== -1 ? Number(rValues[i][rarityIdx]) : 0;
        const uVal = urlIdx !== -1 ? rValues[i][urlIdx] : '';
        if (rVal && uVal) {
          rarityUrlMap[rVal] = String(uVal).trim();
        }
      }
    }

    const tValues = getSafeSheetValues('tag');
    const tagUrlMap = {};
    if (tValues.length >= 2) {
      const tHeaders = tValues[0].map(h => String(h).trim());
      const tagIdx = tHeaders.findIndex(h => ['tag', 'tag_name'].includes(h));
      const tagUrlIdx = tHeaders.indexOf('tag_url');
      if (tagIdx !== -1 && tagUrlIdx !== -1) {
        for (let i = 1; i < tValues.length; i++) {
          const tVal = tValues[i][tagIdx];
          const uVal = tValues[i][tagUrlIdx];
          if (tVal) tagUrlMap[String(tVal).trim()] = uVal ? String(uVal).trim() : '';
        }
      }
    }

    const uValues = getSafeSheetValues('unit');
    const unitTagMap = {};
    const unitInfoMap = {};
    if (uValues.length >= 2) {
      const uHeaders = uValues[0].map(h => String(h).trim());
      const uNameIdx = uHeaders.findIndex(h => ['unit', 'unit_name'].includes(h));
      const uPriorityIdx = uHeaders.indexOf('priority');
      const uTagNameIdx = uHeaders.indexOf('tag_name');
      for (let i = 1; i < uValues.length; i++) {
        const uName = uValues[i][uNameIdx !== -1 ? uNameIdx : 0];
        const uPrio = uPriorityIdx !== -1 ? Number(uValues[i][uPriorityIdx]) || (i * 10) : (i * 10);
        const uTagName = uTagNameIdx !== -1 ? uValues[i][uTagNameIdx] : '';
        if (uName) {
          const trimmedUnitName = String(uName).trim();
          const tagName = uTagName ? String(uTagName).trim() : '';
          const tagUrl = tagUrlMap[tagName] || '';
          unitTagMap[trimmedUnitName] = tagUrl;
          unitInfoMap[trimmedUnitName] = {
            unitName: trimmedUnitName,
            priority: uPrio,
            tagUrl: tagUrl
          };
        }
      }
    }

    const charSheet = ss.getSheetByName('character');
    if (!charSheet) {
      throw new Error('「character」シートが見つかりませんでした。');
    }

    const charValues = charSheet.getDataRange().getValues();
    if (charValues.length < 2) return { characters: [], globalIcons: globalIcons, rarityUrlMap: rarityUrlMap, unitInfoMap: unitInfoMap };

    const headers = charValues[0].map(h => String(h).trim());
    const findIdx = (names) => headers.findIndex(h => names.includes(h));

    const unitIdx = findIdx(['unit', 'ユニット']);
    const iconIdx = findIdx(['icon', 'アイコン']);
    const nameIdx = findIdx(['character', 'キャラクター名', 'character_name', '名前']);
    const plateIdx = findIdx(['character_plate', 'characterplate', 'char_plate']);
    const releaseFlagIdx = findIdx(['release_flag', 'releaseflag', 'release']);

    const characterList = [];
    const charIdMap = {};

    charValues.slice(1).forEach(row => {
      if (row[0] === "" || row[0] === null || row[0] === undefined) return;

      if (releaseFlagIdx !== -1) {
        const flagVal = row[releaseFlagIdx];
        if (flagVal === false || String(flagVal).trim().toLowerCase() === 'false' || flagVal === 0) {
          return;
        }
      }

      const characterName = nameIdx !== -1 ? String(row[nameIdx]).trim() : '';
      if (!characterName) return;

      const charId = Number(row[0]) || 9999;
      charIdMap[characterName] = charId;

      const rawUnitStr = (unitIdx !== -1 && row[unitIdx]) ? String(row[unitIdx]) : '';
      const rawUnits = rawUnitStr.split(/[,、]/).map(u => u.trim()).filter(u => u.length > 0);
      const unitTagUrls = rawUnits.map(u => unitTagMap[u]).filter(url => Boolean(url));

      characterList.push({
        id: charId,
        character: characterName,
        icon: iconIdx !== -1 ? row[iconIdx] : '',
        character_plate: plateIdx !== -1 ? row[plateIdx] : '',
        unit_tag_urls: unitTagUrls,
        cards: []
      });
    });

    characterList.sort((a, b) => a.id - b.id);

    const holomenSheet = ss.getSheetByName('ホロメン') || ss.getSheetByName('character_cards');
    const holomenCardsMap = {};
    if (holomenSheet) {
      const hValues = holomenSheet.getDataRange().getValues();
      if (hValues.length >= 2) {
        const hHeaders = hValues[0].map(h => String(h).trim());
        const findHIdx = (names) => hHeaders.findIndex(h => names.includes(h));

        const cNameIdx = findHIdx(['character', 'キャラクター名', 'character_name', '名前']);
        const iconIdx = findHIdx(['icon', 'アイコン']);
        const imageIdx = findHIdx(['image', '画像', 'メイン画像']);
        const aliasIdx = findHIdx(['alias', '二つ名', 'メンバー名']);
        const rarityIdx = findHIdx(['rarity', 'レアリティ']);
        const releaseAtIdx = findHIdx(['release_at', 'releaseat', '実装日']);
        const typeUrlIdx = findHIdx(['type_url', 'typeurl']);
        const typeIdx = findHIdx(['type', 'タイプ']);
        const hUnitIdx = findHIdx(['unit', 'ユニット', 'unit_name']);
        const pcPlateIdx = findHIdx(['name_plate_url', 'nameplate_url']);
        const mobilePlateIdx = findHIdx(['mobile_name_plate', 'mobilenameplate', 'mobile_plate']);

        for (let i = 1; i < hValues.length; i++) {
          const row = hValues[i];
          if (row[0] === "" || row[0] === null || row[0] === undefined) continue;

          const cName = cNameIdx !== -1 ? String(row[cNameIdx] || '').trim() : '';
          if (!cName) continue;

          if (!holomenCardsMap[cName]) holomenCardsMap[cName] = [];

          const typeStr = typeIdx !== -1 ? String(row[typeIdx] || '').trim() : '';
          const typeUrl = (typeUrlIdx !== -1 && row[typeUrlIdx]) ? row[typeUrlIdx] : (tagUrlMap[typeStr] || '');

          const rawUnitStr = (hUnitIdx !== -1 && row[hUnitIdx]) ? String(row[hUnitIdx]) : '';
          const rawUnits = rawUnitStr.split(/[,、]/).map(u => u.trim()).filter(u => u.length > 0);
          const cardUnitTagUrls = rawUnits.map(u => unitTagMap[u]).filter(url => Boolean(url));

          const pcPlate = pcPlateIdx !== -1 ? row[pcPlateIdx] : '';
          const mobilePlate = mobilePlateIdx !== -1 ? row[mobilePlateIdx] : '';
          const npBg = mobilePlate || pcPlate;

          holomenCardsMap[cName].push({
            id: Number(row[0]) || i,
            icon: iconIdx !== -1 ? row[iconIdx] : '',
            image: imageIdx !== -1 ? row[imageIdx] : '',
            alias: aliasIdx !== -1 ? String(row[aliasIdx] || '').trim() : '',
            character: cName,
            char_id: charIdMap[cName] || 9999,
            rarity: rarityIdx !== -1 ? Number(row[rarityIdx]) || 0 : 0,
            release_at: releaseAtIdx !== -1 ? Number(String(row[releaseAtIdx]).replace(/\D/g, '')) || 0 : 0,
            type_url: typeUrl,
            units: rawUnits,
            unit_tag_urls: cardUnitTagUrls,
            name_plate_bg: npBg
          });
        }
      }
    }

    characterList.forEach(charObj => {
      const cards = holomenCardsMap[charObj.character] || [];
      cards.sort((a, b) => {
        if (b.rarity !== a.rarity) return b.rarity - a.rarity;
        if (b.release_at !== a.release_at) return b.release_at - a.release_at;
        return a.id - b.id;
      });
      charObj.cards = cards;
    });

    return {
      characters: characterList,
      globalIcons: globalIcons,
      rarityUrlMap: rarityUrlMap,
      unitInfoMap: unitInfoMap
    };
  } catch (err) {
    throw new Error('サーバー処理エラー: ' + err.message);
  }
}

/**
 * モーダル表示用：特定のキャラクター詳細データ取得
 */
function getCharacterDetailsByName(charName) {
  try {
    const ss = getTargetSpreadsheet();

    const gValues = getSafeSheetValues('global_icon');
    const globalIcons = {};
    if (gValues.length >= 2) {
      const gHeaders = gValues[0].map(h => String(h).trim());
      const nameIdx = gHeaders.findIndex(h => ['name', 'icon_name', 'type', 'key'].includes(h));
      const urlIdx = gHeaders.indexOf('icon_url');
      for (let i = 1; i < gValues.length; i++) {
        const k = nameIdx !== -1 ? String(gValues[i][nameIdx]).trim() : String(gValues[i][0]).trim();
        const u = urlIdx !== -1 ? gValues[i][urlIdx] : gValues[i][1];
        if (k) globalIcons[k] = u;
      }
    }

    const tValues = getSafeSheetValues('tag');
    const tagUrlMap = {};
    if (tValues.length >= 2) {
      const tHeaders = tValues[0].map(h => String(h).trim());
      const tagIdx = tHeaders.findIndex(h => ['tag', 'tag_name'].includes(h));
      const tagUrlIdx = tHeaders.indexOf('tag_url');
      if (tagIdx !== -1 && tagUrlIdx !== -1) {
        for (let i = 1; i < tValues.length; i++) {
          const tVal = tValues[i][tagIdx];
          const uVal = tValues[i][tagUrlIdx];
          if (tVal) tagUrlMap[String(tVal).trim()] = uVal ? String(uVal).trim() : '';
        }
      }
    }

    const uValues = getSafeSheetValues('unit');
    const unitTagMap = {};
    if (uValues.length >= 2) {
      const uHeaders = uValues[0].map(h => String(h).trim());
      const uNameIdx = uHeaders.findIndex(h => ['unit', 'unit_name'].includes(h));
      const uTagNameIdx = uHeaders.indexOf('tag_name');
      for (let i = 1; i < uValues.length; i++) {
        const uName = uValues[i][uNameIdx !== -1 ? uNameIdx : 0];
        const uTagName = uTagNameIdx !== -1 ? uValues[i][uTagNameIdx] : '';
        if (uName) {
          const tagName = uTagName ? String(uTagName).trim() : '';
          unitTagMap[String(uName).trim()] = tagUrlMap[tagName] || '';
        }
      }
    }

    const charSheet = ss.getSheetByName('ホロメン') || ss.getSheetByName('character');
    if (!charSheet) {
      throw new Error('「ホロメン」または「character」シートが存在しません。');
    }

    const values = charSheet.getDataRange().getValues();
    if (values.length < 2) return { details: [], globalIcons: globalIcons, tagUrlMap: tagUrlMap };

    const headers = values[0].map(h => String(h).trim());
    const nameIdx = headers.findIndex(h => ['character', 'キャラクター名', 'character_name', '名前'].includes(h));

    if (nameIdx === -1) {
      throw new Error('シートの1行目に「character」ヘッダーが見つかりませんでした。');
    }

    const matchedRows = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[0] === "" || row[0] === null || row[0] === undefined) continue;

      if (String(row[nameIdx]).trim() === String(charName).trim()) {
        const rowObj = {};
        headers.forEach((h, idx) => {
          if (h) rowObj[h] = row[idx];
        });

        const rawUnits = rowObj['unit'] ? String(rowObj['unit']).split(/[,、]/).map(u => u.trim()) : [];
        rowObj['unit_tag_urls'] = rawUnits.map(u => unitTagMap[u]).filter(url => Boolean(url));

        const typeStr = rowObj['type'] ? String(rowObj['type']).trim() : '';
        rowObj['type_tag_url'] = tagUrlMap[typeStr] || '';

        matchedRows.push(rowObj);
      }
    }

    matchedRows.sort((a, b) => {
      const rarityA = Number(a['rarity']) || 0;
      const rarityB = Number(b['rarity']) || 0;
      if (rarityA !== rarityB) return rarityB - rarityA;

      const releaseA = Number(String(a['release_at'] || a['releaseat']).replace(/\D/g, '')) || 0;
      const releaseB = Number(String(b['release_at'] || b['releaseat']).replace(/\D/g, '')) || 0;
      if (releaseA !== releaseB) return releaseB - releaseA;

      const idA = Number(a['id']) || 0;
      const idB = Number(b['id']) || 0;
      return idA - idB;
    });

    return {
      details: matchedRows,
      globalIcons: globalIcons,
      tagUrlMap: tagUrlMap
    };
  } catch (err) {
    throw new Error('サーバー処理エラー: ' + err.message);
  }
}

/**
 * 更新履歴シートからデータを取得（2行目以降）
 * 最新データが下に来るように配列を整理
 */
function getUpdateHistory() {
  try {
    const ss = getTargetSpreadsheet();
    const sheet = ss.getSheetByName('更新履歴') || ss.getSheetByName('history') || ss.getSheetByName('log');
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const historyList = [];
    // 2行目以降を順に取得（上が古い、下が最新）
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[0] === "" || row[0] === null || row[0] === undefined) continue;

      const dateVal = row[0] ? formatDateString(row[0]) : '';
      const textVal = row[1] ? String(row[1]).trim() : String(row[0]).trim();

      historyList.push({
        date: dateVal,
        content: textVal
      });
    }

    return historyList;
  } catch (err) {
    console.error('更新履歴取得エラー:', err);
    return [];
  }
}
