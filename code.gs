/**
 * TOBO Quiz Backend (Google Apps Script) - High Performance Edition
 * พัฒนาโดย: 
 * ครูมิน อลีมิน เจ๊ะอีซอ โรงเรียนบ้านเจ๊ะยอ
 * ครูยี แวอัสรี แวมายิ โรงเรียนบ้านปูโป๊ะ
 * สพป.นราธิวาส เขต 2
 */

// ใช้ API Key ตามที่คุณครูระบุ
const API_KEY = "AQ.Ab8RN6LDTEtH_wxPf6xFIjdUDZkuoe1ReoudA3AoTrKy56HzFA";

// 1. ฟังก์ชันบังคับขอสิทธิ์
function forceAuth() {
  UrlFetchApp.fetch("https://www.google.com");
}

// 2. ฟังก์ชันตั้งค่าฐานข้อมูล
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let gameSheet = ss.getSheetByName('Games');
  if (!gameSheet) {
    gameSheet = ss.insertSheet('Games');
    gameSheet.appendRow(['Game PIN', 'Topic', 'Grade', 'Status', 'CreatedAt', 'Questions_JSON']);
    gameSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#d9e2f3');
  }

  let playerSheet = ss.getSheetByName('Players');
  if (!playerSheet) {
    playerSheet = ss.insertSheet('Players');
    playerSheet.appendRow(['Timestamp', 'Game PIN', 'Player Name', 'Score', 'Rank']);
    playerSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e2efd9');
  }

  let dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) {
    dashSheet = ss.insertSheet('Dashboard');
    dashSheet.appendRow(['Total Games Created', 'Total Players Joined']);
    dashSheet.appendRow([0, 0]);
    dashSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#fff2cc');
  }
}

// 3. ฟังก์ชันหน้าเว็บ
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('TOBO Quiz - สื่อการสอนยุคใหม่')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 4. ฟังก์ชัน AI (รองรับ Key แบบ AQ... ของครูทราย)
function generateQuizFromAI(topic, gradeLevel, questionCount) {
  try {
    const prompt = `สร้างคำถามปรนัย ${questionCount} ข้อเรื่อง "${topic}" ระดับ "${gradeLevel}"
    เงื่อนไขสำคัญ:
    1. คำถามต้อง "สั้น กระชับ เข้าใจง่าย"
    2. ตัวเลือก 4 ข้อ สั้นๆ
    3. ต้องมี "คำอธิบาย (explanation)" สั้นๆ 1 ประโยค ว่าทำไมข้อนี้ถึงถูก
    ตอบกลับเป็น JSON Array โครงสร้าง:
    [{"question": "คำถาม?", "options": ["A", "B", "C", "D"], "correctIndex": 0, "explanation": "อธิบายสั้นๆ"}]`;

    let url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent`;
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: "application/json" }
    };
    
    let options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    // ตรวจสอบและส่ง Key เข้าไปในระบบตามวิธีที่ถูกต้อง
    if(API_KEY.startsWith("AIza")) {
      url += `?key=${API_KEY}`;
    } else {
      // ส่ง Key รุ่น AQ... เข้าทาง Header (OAuth2) ตามที่ Google เรียกร้อง
      options.headers = { "Authorization": "Bearer " + API_KEY };
    }
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.error) throw new Error(data.error.message);
    if (!data.candidates || data.candidates.length === 0) throw new Error("AI ไม่ส่งข้อมูลกลับมา");
    
    let rawText = data.candidates[0].content.parts[0].text;
    rawText = rawText.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(rawText);
  } catch (err) {
    throw new Error("AI Error: " + err.message);
  }
}

// 5. API จัดการข้อมูลเกม (CacheService - เร็วพิเศษ)
function api(action, payload) {
  const cache = CacheService.getScriptCache();
  try {
    if (action === 'getDashboardStats') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dashSheet = ss.getSheetByName('Dashboard');
      if(!dashSheet) return { success: true, totalGames: 0, totalPlayers: 0 };
      return {
        success: true,
        totalGames: dashSheet.getRange('A2').getValue() || 0,
        totalPlayers: dashSheet.getRange('B2').getValue() || 0
      };
    }
    
    if (action === 'createGame') {
      const pin = Math.floor(1000 + Math.random() * 9000).toString();
      const newGame = {
        pin: pin,
        status: 'lobby',
        currentQuestionIndex: 0,
        timeLimit: payload.timeLimit || 15,
        players: {}, 
        questions: payload.questions,
        topic: payload.topic,
        gradeLevel: payload.gradeLevel,
        createdAt: Date.now()
      };
      cache.put('GAME_' + pin, JSON.stringify(newGame), 21600);
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dashS = ss.getSheetByName('Dashboard');
      if(dashS) dashS.getRange('A2').setValue(Number(dashS.getRange('A2').getValue() || 0) + 1);
      
      return { success: true, pin: pin };
    }

    if (action === 'joinGame') {
      let gameDataStr = cache.get('GAME_' + payload.pin);
      if (!gameDataStr) throw new Error("ไม่พบห้องเกมรหัสนี้ หรือเกมจบไปแล้ว");
      
      let gameData = JSON.parse(gameDataStr);
      if (gameData.status !== 'lobby') throw new Error("เกมนี้เริ่มไปแล้วครับ!");

      const uid = 'p_' + Date.now() + Math.floor(Math.random()*1000);
      gameData.players[uid] = { name: payload.name, score: 0, lastAnswer: null, answerTime: null };
      cache.put('GAME_' + payload.pin, JSON.stringify(gameData), 21600);
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dashP = ss.getSheetByName('Dashboard');
      if(dashP) dashP.getRange('B2').setValue(Number(dashP.getRange('B2').getValue() || 0) + 1);

      return { success: true, uid: uid, game: gameData };
    }

    if (action === 'getGameState') {
      const currentData = cache.get('GAME_' + payload.pin);
      if (!currentData) throw new Error("ไม่พบข้อมูลห้อง หรือห้องถูกปิดแล้ว");
      return { success: true, game: JSON.parse(currentData) };
    }

    if (action === 'updateGameState') {
      let hostGameStr = cache.get('GAME_' + payload.pin);
      if (hostGameStr) {
        let hostGame = JSON.parse(hostGameStr);
        Object.assign(hostGame, payload.updates);
        cache.put('GAME_' + payload.pin, JSON.stringify(hostGame), 21600);
        return { success: true };
      }
      throw new Error("ห้องเกมสูญหาย หรือหมดเวลาแล้ว");
    }

    if (action === 'submitAnswer') {
      let pGameStr = cache.get('GAME_' + payload.pin);
      if (pGameStr) {
        let pGame = JSON.parse(pGameStr);
        if (pGame.players[payload.uid]) {
          pGame.players[payload.uid].lastAnswer = payload.answerIndex;
          pGame.players[payload.uid].score = payload.score;
          pGame.players[payload.uid].answerTime = Date.now();
          cache.put('GAME_' + payload.pin, JSON.stringify(pGame), 21600);
          return { success: true };
        }
      }
      throw new Error("ส่งคำตอบไม่สำเร็จ");
    }

    if (action === 'endGameAndSave') {
      let finalGameStr = cache.get('GAME_' + payload.pin);
      if (finalGameStr) {
        let finalGame = JSON.parse(finalGameStr);
        const activeSs = SpreadsheetApp.getActiveSpreadsheet();
        activeSs.getSheetByName('Games').appendRow([
          finalGame.pin, finalGame.topic, finalGame.gradeLevel, 'Completed', new Date(), JSON.stringify(finalGame.questions)
        ]);

        const playersArr = Object.values(finalGame.players).sort((a,b) => b.score - a.score);
        const pSheet = activeSs.getSheetByName('Players');
        const timestamp = new Date();
        
        let pDataToSave = [];
        playersArr.forEach((p, idx) => {
           pDataToSave.push([timestamp, finalGame.pin, p.name, p.score, idx + 1]);
        });
        if (pDataToSave.length > 0) pSheet.getRange(pSheet.getLastRow() + 1, 1, pDataToSave.length, 5).setValues(pDataToSave);
        
        cache.remove('GAME_' + payload.pin);
        return { success: true };
      }
      return { success: false };
    }

  } catch (error) {
    return { success: false, error: error.message };
  }
}
