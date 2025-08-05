import os
import datetime
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def transfer_yahoo_news_from_source_sheet():
    SOURCE_SPREADSHEET_ID = '1ELh95L385GfNcJahAx1mUH4SZBHtKImBp_wAAsQALkM'
    DESTINATION_SPREADSHEET_ID = '1ELh95L385GfNcJahAx1mUH4SZBHtKImBp_wAAsQALkM'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    try:
        creds_json = os.environ.get('GCP_SA_KEY')
        if not creds_json:
            with open('key.json', 'r') as f:
                creds_info = json.load(f)
        else:
            creds_info = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
    except Exception as e:
        print(f"エラー: Google Sheets APIの認証に失敗しました。詳細: {e}")
        return

    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))) 
    yesterday = today - datetime.timedelta(days=1)
    start_time = yesterday.replace(hour=15, minute=0, second=0, microsecond=0)
    end_time = today.replace(hour=14, minute=59, second=59, microsecond=0)
    destination_sheet_name = today.strftime('%y%m%d')

    print(f"出力先シート名: {destination_sheet_name}")
    print(f"期間: {start_time.strftime('%Y/%m/%d %H:%M:%S')} 〜 {end_time.strftime('%Y/%m/%d %H:%M:%S')}")

    existing_data_in_destination = []
    header_exists = False
    try:
        spreadsheet_info = service.spreadsheets().get(spreadsheetId=DESTINATION_SPREADSHEET_ID).execute()
        sheets = spreadsheet_info.get('sheets', [])
        sheet_exists = any(sheet['properties']['title'] == destination_sheet_name for sheet in sheets)
        if not sheet_exists:
            print(f"出力先スプレッドシートに新しいシート「{destination_sheet_name}」を作成します。")
            body = {'requests': [{'addSheet': {'properties': {'title': destination_sheet_name}}}]}
            service.spreadsheets().batchUpdate(spreadsheetId=DESTINATION_SPREADSHEET_ID, body=body).execute()
            print(f"新しいシート「{destination_sheet_name}」を作成しました。")
        destination_sheet_range = f"'{destination_sheet_name}'!A:L"
        result = service.spreadsheets().values().get(spreadsheetId=DESTINATION_SPREADSHEET_ID, range=destination_sheet_range).execute()
        existing_data_in_destination = result.get('values', [])
    except HttpError as e:
        print(f"エラー: 出力先スプレッドシートへのアクセスに失敗しました。詳細: {e}")
        return
    except Exception as e:
        print(f"エラー: 不明なエラーが発生しました。詳細: {e}")
        return

    existing_urls_in_destination = set()
    if existing_data_in_destination and existing_data_in_destination[0][0] == 'ソース':
        header_exists = True
        for row in existing_data_in_destination[1:]:
            if len(row) > 2 and row[2]:
                existing_urls_in_destination.add(row[2])

    print(f"出力先シートに既存のニュースが {len(existing_urls_in_destination)} 件あります（URLで重複を判定）。")

    source_sheet_name = 'Yahoo'
    try:
        source_sheet_range = f"'{source_sheet_name}'!A:D"
        result = service.spreadsheets().values().get(spreadsheetId=SOURCE_SPREADSHEET_ID, range=source_sheet_range).execute()
        data = result.get('values', [])
    except Exception as e:
        print(f"エラー: コピー元シート「{source_sheet_name}」にアクセスできませんでした。詳細: {e}")
        return

    if not data:
        print(f"エラー: コピー元シート「{source_sheet_name}」にデータがありません。")
        return

    print(f"シート「{source_sheet_name}」から {len(data) - 1} 件のニュースを読み込みました（ヘッダーを除く）。")
    
    collected_news = []
    for i, row in enumerate(data):
        if i == 0:
            continue
        try:
            title, url, post_date_raw, source = row
            post_date = None
            if isinstance(post_date_raw, str):
                try:
                    post_date_without_year = datetime.datetime.strptime(post_date_raw, '%m/%d %H:%M')
                    post_date = post_date_without_year.replace(year=today.year)
                except ValueError:
                    try:
                        post_date = datetime.datetime.strptime(post_date_raw, '%Y/%m/%d %H:%M:%S')
                    except ValueError:
                        pass
            if post_date:
                post_date = post_date.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=9)))
                if start_time <= post_date <= end_time and url not in existing_urls_in_destination:
                    new_row = [source_sheet_name, title, url, post_date.strftime('%Y/%m/%d'), source]
                    collected_news.append(new_row)
        except Exception as e:
            print(f"警告: 行 {i+1} でエラー: {e}")
            continue

    if not collected_news:
        print("期間内に新しいニュースは見つかりませんでした。")
        return

    print(f"新規に追記するニュースの数: {len(collected_news)}")

    data_to_append = []
    for i, row in enumerate(collected_news):
        row_data = row + [''] * 4
        k_value = row[1][:20]
        l_value = i + 1
        row_data += ['', k_value, l_value]
        data_to_append.append(row_data)

    if not header_exists:
        header_row = ['ソース', 'タイトル', 'URL', '投稿日', '引用元', 'コメント数', 'ポジネガ', 'カテゴリー', '有料記事', 'J列', 'K列', 'L列']
        service.spreadsheets().values().append(spreadsheetId=DESTINATION_SPREADSHEET_ID, range=f"'{destination_sheet_name}'!A1", valueInputOption='USER_ENTERED', body={'values': [header_row]}).execute()
        print(f"シート「{destination_sheet_name}」にヘッダーを追加しました。")

    service.spreadsheets().values().append(
        spreadsheetId=DESTINATION_SPREADSHEET_ID,
        range=f"'{destination_sheet_name}'!A:L",
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': data_to_append}
    ).execute()
    print(f"スプレッドシートに {len(data_to_append)} 件を追記しました。")

if __name__ == '__main__':
    transfer_yahoo_news_from_source_sheet()
