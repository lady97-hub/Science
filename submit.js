/**
 * ==============================================================================
 * 구글 앱스 스크립트(Google Apps Script, GAS) 연동 스크립트
 * ==============================================================================
 * [선생님을 위한 구글 스프레드시트 1분 연동 가이드]
 * 1. Google Drive에서 새 스프레드시트를 만듭니다.
 * 2. 상단 메뉴 [확장 프로그램] -> [Apps Script]를 클릭합니다.
 * 3. 기존 코드를 지우고, 아래 [GAS 서버 코드]를 그대로 복사하여 붙여넣고 저장(Ctrl+S)합니다.
 * 4. 오른쪽 위 [배포] -> [새 배포] 클릭
 *    - 유형 선택: 웹 앱 (Web App)
 *    - 설명: 초등 6학년 디지털 시민성 퀴즈 제출
 *    - 다음 사용자 권한으로 실행: '나(내 계정)'
 *    - 액세스 권한이 있는 사용자: '모든 사용자(Anyone)' (※ 학생 로그인이 필요 없도록 필수 설정!)
 * 5. [배포]를 누른 뒤 생성된 '웹 앱 URL'을 복사하여 아래 GAS_WEB_APP_URL 에 붙여넣으세요!
 *
 * ------------------------------------------------------------------------------
 * [GAS 서버 코드 (Apps Script 에디터에 붙여넣을 코드)]
 * ------------------------------------------------------------------------------
 * function doPost(e) {
 *   try {
 *     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 *     var data = JSON.parse(e.postData.contents);
 *     
 *     // 시트가 비어있으면 헤더(제목 행) 자동 생성
 *     if (sheet.getLastRow() === 0) {
 *       var headers = ["제출 시각", "취득 점수", "총 문항 수", "달성률(%)"];
 *       for (var i = 1; i <= 10; i++) {
 *         headers.push("Q" + i + " 정오답");
 *       }
 *       sheet.appendRow(headers);
 *       sheet.getRange(1, 1, 1, headers.length)
 *            .setBackground("#4F46E5")
 *            .setFontColor("#FFFFFF")
 *            .setFontWeight("bold");
 *     }
 *     
 *     // 행 데이터 구성 (개인정보는 절대 포함되지 않음)
 *     var row = [
 *       data.submittedAt,
 *       data.score,
 *       data.totalQuestions,
 *       Math.round((data.score / data.totalQuestions) * 100) + "%"
 *     ];
 *     
 *     // Q1~Q10 정오답 결과 추가 (O 또는 X)
 *     if (data.questionResults && data.questionResults.length > 0) {
 *       for (var j = 0; j < data.questionResults.length; j++) {
 *         row.push(data.questionResults[j].isCorrect ? "O (정답)" : "X (오답)");
 *       }
 *     }
 *     
 *     sheet.appendRow(row);
 *     
 *     return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
 *                          .setMimeType(ContentService.MimeType.JSON);
 *   } catch (err) {
 *     return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
 *                          .setMimeType(ContentService.MimeType.JSON);
 *   }
 * }
 * ==============================================================================
 */

// ★ 배포한 Google Apps Script 웹 앱 URL을 여기에 붙여넣으세요.
// 예시: "https://script.google.com/macros/s/AKfycbx.../exec"
const GAS_WEB_APP_URL = "YOUR_GAS_WEB_APP_URL_HERE";

/**
 * 퀴즈 응답 결과를 구글 앱스 스크립트로 전송하는 함수
 * @param {Object} quizSummary - { score, totalQuestions, questionResults }
 * @param {Function} callback - 상태 메시지 업데이트 콜백 (status, message)
 */
async function submitQuizResultToSheet(quizSummary, callback) {
  // 현재 제출 일시 포맷 (예: 2026. 10. 24. 오후 1:30:15)
  const now = new Date();
  const formattedTime = now.toLocaleDateString('ko-KR', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    weekday: 'short'
  }) + " " + now.toLocaleTimeString('ko-KR');

  // 전송할 페이로드 구성: 완전 익명 보장 (개인정보 일체 제외)
  const payload = {
    submittedAt: formattedTime,
    score: quizSummary.score,
    totalQuestions: quizSummary.totalQuestions,
    questionResults: quizSummary.questionResults.map(item => ({
      questionId: item.id,
      category: item.category,
      isCorrect: item.isCorrect,
      selectedAnswerIndex: item.selectedAnswerIndex + 1, // 1~4번
      correctAnswerIndex: item.correctAnswerIndex + 1
    }))
  };

  // URL이 기본 플레이스홀더이거나 비어있는 경우 시뮬레이션 모드로 친절하게 안내
  if (!GAS_WEB_APP_URL || GAS_WEB_APP_URL === "YOUR_GAS_WEB_APP_URL_HERE") {
    console.log("[익명 퀴즈 제출 데이터 (테스트 모드)]", payload);
    if (callback) {
      callback({
        status: "demo",
        message: "연동 안내: submit.js에 구글 시트 Apps Script URL을 넣으면 실시간 자동 기록됩니다. (지금은 브라우저 테스트 완료 상태입니다)",
        payload: payload
      });
    }
    return;
  }

  // 실제 구글 스프레드시트 Apps Script로 비동기 전송
  try {
    if (callback) {
      callback({
        status: "loading",
        message: "구글 스프레드시트로 안전하게 응답을 전송하고 있습니다..."
      });
    }

    // Google Apps Script는 브라우저 CORS 보안 특성상 no-cors 모드로 전송할 때 가장 안정적으로 기록됩니다.
    await fetch(GAS_WEB_APP_URL, {
      method: "POST",
      mode: "no-cors",
      cache: "no-cache",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    if (callback) {
      callback({
        status: "success",
        message: "구글 스프레드시트에 응답이 성공적으로 기록되었습니다! (개인정보는 수집되지 않았습니다)",
        payload: payload
      });
    }
  } catch (error) {
    console.error("구글 시트 전송 실패:", error);
    if (callback) {
      callback({
        status: "error",
        message: "구글 시트 전송 중 일시적인 연결 오류가 발생했습니다. (학생의 퀴즈 결과는 화면에서 정상 확인할 수 있습니다)",
        payload: payload
      });
    }
  }
}
