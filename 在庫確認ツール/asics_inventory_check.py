import eel
import pandas as pd
import requests
from bs4 import BeautifulSoup
import traceback

# 靴のサイズの変更
size_dict = {
    '22.5':'US4/JP22.5',
    '23':'US4.5/JP23.0',
    '23.5':'US5/JP23.5',
    '24':'US5.5/JP24.0',
    '24.5':'US6/JP24.5',
    '25':'US6.5/JP25',
    '25.5':'US7.5/JP25.5',
    '26':'US8/JP26',
    '26.5':'US8.5/JP26.5',
    '27':'US9/JP27',
    '27.5':'US9.5/JP27.5',
    '28':'US10/JP28',
    '28.5':'US11/JP28.5',
    '29':'US11.5/JP29',
    '29.5':'US12/JP29.5',
    '30':'US12.5/JP30',
    '30.5':'US13/JP30.5',
    '31':'US14/JP31',
    '32':'US15/JP32',
    '33':'US16/JP33'
}

class MyScraping():

    def __init__(self, items_id, items_size):
        self.items_id = items_id
        self.items_size = items_size

    # 靴のサイズの変更(日本→記載)
    def change_shoose_size(self, shoose_jp_size):
        return size_dict[shoose_jp_size]

    # listの大きさをそろえる
    def create_same_size_list(self, size_list, list, append_word):
        return list + [append_word] * (len(size_list) - len(list))

    #一覧ページから情報を取り出す
    def get_item_detail(self, item_df):

        i = 1
        for index, row in item_df.iterrows():
            eel.view_log_js(str(i) + "/" + str(len(item_df)) + " 商品目")

            now_size_list = row[0].replace("Size=", "").split(",")

            try:
                url = 'https://www.asics.com/jp/ja-jp/' + row[1]
                res = requests.get(url)
                soup = BeautifulSoup(res.text, 'lxml')

                # 靴のサイズ
                now_items_size_list_len = len(self.items_size)
                size_list = soup.find_all(class_ = "variants__item--size")
                for size in size_list:
                    data_instock = size["data-instock"]
                    if data_instock == "false":

                        try:
                            size = self.change_shoose_size(size.get_text(strip=True))
                            size_replace_H_list = size.split("/")
                            size_replace_H = size_replace_H_list[0].replace(".5", "H") + "/" + size_replace_H_list[1]

                            if size in now_size_list:
                                self.items_size.append("Size=" + size)
                                continue
                            if size_replace_H in now_size_list:
                                self.items_size.append("Size=" + size_replace_H)
                                continue
                        except KeyError:
                            continue
                
                after_items_size_list_len = len(self.items_size)
                if now_items_size_list_len == after_items_size_list_len:
                    i += 1
                    continue

                #listの格納
                self.items_id.append(row[1])

                #listのサイズの調整
                self.items_id = self.create_same_size_list(self.items_size, self.items_id, "")

                i += 1

            except Exception:
                t = traceback.format_exc()
                eel.view_log_js(t)
                eel.view_log_js("予期せぬエラーが発生しました。")
                continue

            except KeyboardInterrupt:
                eel.view_log_js("プログラムを強制的に終了しました")
                return self.items_id, self.items_size

        return self.items_id, self.items_size

def main():

    try:
        myscraping = MyScraping(items_id=[], items_size=[])

        dataframe = pd.read_excel(eel.file_name()()).dropna(subset=['RelationshipDetails','CustomLabel'])
        items_id, items_size = myscraping.get_item_detail(dataframe)

        eel.view_log_js('\n各商品のサイズの取得が終わりました。')

    finally:
        eel.view_log_js('\nファイルの作成をします')

        item_change_df = pd.DataFrame([
            items_size,
            items_id
        ]).T

        if item_change_df.empty:
            item_change_df = pd.DataFrame(["在庫切れの商品はありませんでした"])
            eel.view_log_js('\n在庫切れの商品はありませんでした。')

        item_change_df.to_excel(eel.file_name()().replace(".xlsx", "_在庫確認.xlsx"), header=False, index=False)

        eel.view_log_js('\nファイルの作成が完了しました')
