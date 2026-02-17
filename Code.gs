function onFormSubmit(e) {
  try {
    // --- 設定項目 ---
    const calendarId = 'example.calendar.google.com';
    const adminEmails = 'example1@gmail.com, example2@gmail.com'; // カンマ区切りで追加
    // ---------------

    // eがundefinedの場合のエラーハンドリング
    if (!e || !e.response) {
      throw new Error('フォーム送信イベントが正しく渡されていません');
    }

    const itemResponse = e.response.getItemResponses();
    const name = itemResponse[0].getResponse(); // お名前
    const dateStr = itemResponse[1].getResponse(); // 練習日
    const startTime = itemResponse[2].getResponse(); // 開始
    const endTime = itemResponse[3].getResponse(); // 終了

    const start = new Date(dateStr + ' ' + startTime);
    const end = new Date(dateStr + ' ' + endTime);
    const calendar = CalendarApp.getCalendarById(calendarId);

    // 指定された時間にすでに入っている予定を取得
    const events = calendar.getEvents(start, end);

    // 予約人数チェック（2人までOK、3人目はNG）
    if (events.length >= 2) {
      const userEmail = e.response.getRespondentEmail();
      const errorSubject = '【予約不可】予約が埋まっています';
      const errorBody = name + ' 様\n\n申し訳ありません。' + dateStr + ' ' + startTime + ' の枠はすでに2名の予約が入っているため、予約できませんでした。別の時間帯で再度お申し込みください。';
      
      if (userEmail) {
        MailApp.sendEmail(userEmail, errorSubject, errorBody);
      }
      MailApp.sendEmail(adminEmails, '【予約失敗】重複のため却下されました', name + 'さんの予約が定員オーバーで失敗しました。');
      
      return;
    }

    // 2名以下なら通常通り登録
    calendar.createEvent('重機練習：' + name, start, end);
    
    // 成功通知メール
    const subject = '【重機練習】予約完了のお知らせ';
    const body = name + ' さんの予約を登録しました。\n日時：' + dateStr + ' ' + startTime + ' 〜 ' + endTime;
    MailApp.sendEmail(adminEmails, subject, body);

  } catch (error) {
    // エラーログを記録
    Logger.log('エラー: ' + error.message);
    // 管理者にエラー通知
    MailApp.sendEmail(adminEmails, '【システムエラー】予約処理に失敗', 'エラー内容: ' + error.message + '\n\nスタック: ' + error.stack);
  }
}
