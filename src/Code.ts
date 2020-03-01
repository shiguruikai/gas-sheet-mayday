import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;

type SheetValue = [string, string, boolean, string]

interface TvProgram {
    title: string;
    startDateTime: string;
    recorded: boolean;
    url: string;
}

interface SearchResult {
    status: string;
    body: {
        "@odata.context": string;
        "@odata.count": number;
        value: [
            {
                title: string;
                cid: string;
                channel_type: string;
                channel_name: string;
                channel_no: number;
                is_channel_4k: boolean;
                channel_logo_url: string;
                channel_logo_alt: string;
                program_date_0_23: string;
                start_hhmmss: string;
                date: string;
                day_of_week: string;
                start_time: string;
                end_time: string;
                air_time: number;
                start_date: {
                    date: string;
                    timezone_type: number;
                    timezone: string;
                };
                is_show_remote_rec_btn_for_list: boolean;
                within_remote_rec_piriod: boolean;
                si_genre: string;
            }
        ];
    };
}

const SHEET_NAME = "main";
const UNTIL_DATE = 5;

function execute() {
    const scriptLock = LockService.getScriptLock();

    if (scriptLock.tryLock(1)) {
        try {
            updateSheet();
        } finally {
            scriptLock.releaseLock();
        }
    }
}

function updateSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME) ?? spreadsheet.insertSheet().setName(SHEET_NAME);

    const currentTvPrograms: TvProgram[] = sheet.getRange("A2:D").getValues().map(it => ({
        title: normalizeTitle(it[0]),
        startDateTime: it[1],
        recorded: it[2],
        url: it[3],
    }));

    // curl -s 'https://tvguide.myjcom.jp/api/mypage/get_searchresult/' -F 'keyword=メーデー！' -F 'channel=546_65406' -F 'offset=0' | jq
    const fetchedTvPrograms: TvProgram[] = fetchTvPrograms("メーデー", "546_65406", 3);

    // タイトルの重複を消すためのマップ
    const map: { [title: string]: TvProgram } = {};

    // 現在のデータをセット
    currentTvPrograms.forEach(it => {
        if (it.title) {
            map[it.title] = it;
        }
    });

    // 検索した番組データをセット
    // 放送日が近い日付をセットしたいため逆順で回す
    fetchedTvPrograms.reverse().forEach(it => {
        const tvp = map[it.title];
        if (tvp) {
            it.recorded = tvp.recorded;
        }
        map[it.title] = it;
    });

    const newValues = objectValues(map).map(it => [it.title, it.startDateTime, it.recorded, it.url]);

    const headValues = [["タイトル", "放送時間", "録画済み", "URL"]];

    // シートをクリア
    sheet.clear();

    // ヘッダーを追加
    sheet
        .getRange(1, 1, headValues.length, headValues[0].length)
        .setValues(headValues)
        .setHorizontalAlignment("center");

    // 値を追加
    sheet.getRange(2, 1, newValues.length, newValues[0].length).setValues(newValues);

    // データ範囲を取得
    const dataRange = sheet.getDataRange();
    const firstRow = dataRange.getRow();
    const lastRow = dataRange.getLastRow();

    // 「放送時間」の日付をフォーマット
    sheet.getRange("B2:B" + lastRow).setNumberFormat("mm月dd日(ddd) hh:mm");

    // 「録画済み」の値をチェックボックス化
    sheet.getRange("C2:C" + lastRow).insertCheckboxes();

    // 交互の背景色
    dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

    // 現在のフィルターを削除
    sheet.getFilter()?.remove();

    // 新たにフィルターを作成
    const filter = dataRange.createFilter();
    // 「放送時間」の近い順にソート
    filter.sort(2, true);
    // 「放送時間」が今日以降のデータだけにフィルタ
    const afterTodayFilter = SpreadsheetApp.newFilterCriteria()
        .whenDateAfter(SpreadsheetApp.RelativeDate.TODAY)
        .build();
    filter.setColumnFilterCriteria(2, afterTodayFilter);

    // 未録画かつ放送時間がUNTIL_DATE日後のデータを取得
    const conditionDate = new Date();
    conditionDate.setDate(conditionDate.getDate() + UNTIL_DATE);
    const unrecordedValues = (dataRange.getValues() as SheetValue[]).filter((value, index) => {
        return !value[2]
            && !sheet.isRowHiddenByFilter(index + firstRow)
            && new Date(value[1]).getTime() < conditionDate.getTime();
    });

    // 未録画番組があれば通知
    if (unrecordedValues.length > 0) {
        notifyToGmail(spreadsheet, unrecordedValues);
        notifyToSlack(spreadsheet, unrecordedValues);
    }
}

function fetchTvPrograms(keyword: string, channel: string | null, maxPage: number): TvProgram[] {
    const tvPrograms: TvProgram[] = [];

    let page = 0;
    let offsetCount = 0;
    while (page++ < maxPage) {
        const searchResult = fetchSearchResult(keyword, channel, offsetCount);
        const totalCount = searchResult.body["@odata.count"];
        const count = searchResult.body.value.length;

        if (offsetCount >= totalCount || totalCount <= 0 || count <= 0) {
            break;
        }

        offsetCount += count;

        searchResult.body.value.forEach(it => {
            const date = toDate(it.start_date.date.trim());
            const startDateTime = Utilities.formatDate(date, "UTC", "yyyy/MM/dd H:m:s");

            tvPrograms.push({
                url: "https://tvguide.myjcom.jp/detail/?eid=" + it.cid,
                title: normalizeTitle(it.title),
                startDateTime: startDateTime,
                recorded: false,
            });
        });

        Utilities.sleep(1000);
    }

    return tvPrograms;
}

function fetchSearchResult(keyword: string, channel: string | null, offset: number): SearchResult {
    const url = "https://tvguide.myjcom.jp/api/mypage/get_searchresult/";

    const formData = {
        keyword: keyword,
        channel: channel,
        offset: offset,
    };

    const options = {
        type: "POST",
        payload: formData,
    };

    const response = UrlFetchApp.fetch(url, options);

    return JSON.parse(response.getContentText()) as SearchResult;
}

function notifyToGmail(spreadsheet: Spreadsheet, unrecordedValues: SheetValue[]) {
    const recipient = PropertiesService.getScriptProperties().getProperty("RECIPIENT_EMAIL");

    if (recipient === null) {
        Logger.log("Cannot get property 'RECIPIENT_EMAIL'.");
        return;
    }

    const subject = "『メーデー！』を録画するのだ！";

    const textBody = `${spreadsheet.getName()} ${spreadsheet.getUrl()}`;

    const htmlBody = unrecordedValues.map(value => `<a href="${value[3]}">${value[0]}</a>`).join("<br>")
        + `<br><a href="${spreadsheet.getUrl()}">${spreadsheet.getName()}</a>`;

    const options = {
        htmlBody: htmlBody
    };

    MailApp.sendEmail(recipient, subject, textBody, options);
}

function notifyToSlack(spreadsheet: Spreadsheet, unrecordedValues: SheetValue[]) {
    const url = PropertiesService.getScriptProperties().getProperty("SLACK_WEB_HOOK_URL");

    if (url === null) {
        Logger.log("Cannot get property 'SLACK_WEB_HOOK_URL'.");
        return;
    }

    const payload = JSON.stringify({
        channel: "#bot",
        username: "GAS bot",
        icon_emoji: ":robot_face:",
        text: "『メーデー！』を録画するのだ！\n"
            + unrecordedValues.map(it => `<${it[3]}|${it[0]}>`).join("\n")
            + `\n<${spreadsheet.getUrl()}|${spreadsheet.getName()}>`
    });

    const response = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: payload,
        muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
        Logger.log("Failed to send notification to Slack. "
            + `responseCode=${responseCode}, contentText=${response.getContentText()}`);
    }
}

function normalizeTitle(title: string) {
    return title
        .replace(/^\s*●.*?●\s*/, "")
        .replace(/\s*\[[^\[\]]*]\s*/g, "")
        .trim();
}

function toDate(dateString: string): Date {
    const [y, M, d, h, m, s] = dateString.split(/[^\d]+/, 6).map(n => Number(n));
    return new Date(Date.UTC(y, M - 1, d, h, m, s));
}

function objectValues<T>(o: { [key: string]: T }): T[] {
    return Object.keys(o).map(it => o[it]);
}
