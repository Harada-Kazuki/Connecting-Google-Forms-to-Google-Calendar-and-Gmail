function onFormSubmit(e) {
  const calendarId = '4c6253d7b6867972c9b884453f8957229acd6d5631b4be2c5917bdb4b035be54@group.calendar.google.com';
  const adminEmails = 'kawaguchigumi.0022@gmail.com, kawaguchigumi.0007@icloud.com, harada.kazuki@kawaguchigumi.co.jp';

  try {
    if (!e || !e.response) {
      throw new Error('フォーム送信イベントが正しく渡されていません');
    }

    const itemResponse = e.response.getItemResponses();
    const name     = itemResponse[0].getResponse();
    const dateStr  = itemResponse[1].getResponse();
    const startTime = itemResponse[2].getResponse();
    const endTime   = itemResponse[3].getResponse();

    const start = parseDateTime(dateStr, startTime);
    const end   = parseDateTime(dateStr, endTime);

    // 時刻の整合性チェック
    if (isNaN(start) || isNaN(end)) {
      throw new Error('日時のパースに失敗しました: ' + dateStr + ' ' + startTime + '-' + endTime);
    }
    if (end <= start) {
      throw new Error('終了時刻が開始時刻以前です');
    }

    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      throw new Error('カレンダーが見つかりません。IDを確認してください。');
    }

    // 厳密な重複チェック
    const events = calendar.getEvents(start, end);
    const overlapping = events.filter(ev =>
      ev.getStartTime() < end && ev.getEndTime() > start
    );

    const userEmail = e.response.getRespondentEmail();

    if (overlapping.length >= 2) {
      const errorBody = `${name} 様\n\n申し訳ありません。${dateStr} ${startTime} の枠はすでに2名の予約が入っているため、予約できませんでした。\n別の時間帯で再度お申し込みください。`;

      if (userEmail) {
        MailApp.sendEmail(userEmail, '【予約不可】重機練習の予約が埋まっています', errorBody);
      }
      MailApp.sendEmail(adminEmails, '【予約失敗】重複のため却下', `${name} さんの予約が定員オーバーで失敗しました。\n日時：${dateStr} ${startTime}〜${endTime}`);
      return;
    }

    // 予約登録
    calendar.createEvent(`重機練習：${name}`, start, end);

    // 成功通知
    const successBody = `${name} さんの予約を登録しました。\n日時：${dateStr} ${startTime} 〜 ${endTime}\n現在の予約数：${overlapping.length + 1} / 2`;
    MailApp.sendEmail(adminEmails, '【重機練習】予約完了のお知らせ', successBody);

    if (userEmail) {
      MailApp.sendEmail(userEmail, '【重機練習】予約完了のお知らせ', `${name} 様\n\n予約が完了しました。\n日時：${dateStr} ${startTime} 〜 ${endTime}\n\nよろしくお願いいたします。`);
    }

  } catch (error) {
    Logger.log('エラー: ' + error.message);
    MailApp.sendEmail(adminEmails, '【システムエラー】予約処理に失敗', `エラー内容: ${error.message}\n\nスタック: ${error.stack}`);
  }
}

function parseDateTime(dateStr, timeStr) {
  const normalized = dateStr.trim().replace(/\//g, '-');
  const dt = new Date(`${normalized}T${timeStr.trim()}:00`);
  return dt;
}
