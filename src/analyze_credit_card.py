import pandas as pd
import numpy as np
import openpyxl
from tkinter import filedialog
import matplotlib.pyplot as plt
import io
from basic_inf import Basic_inf

#表を作成するスクリプトは別ファイルに用意したため、importする。
from depict_graph import Show_graph

#クレジットカード明細を分析するクラス
class Analyze_credit_card():

    #ファイルの読み込み
    def load_csv_file():

        title = '利用明細を選択してください。'
        
        filetype = [('csvファイル','*.csv')]

        Basic_inf.csv_file_path = filedialog.askopenfilename(title= title,filetypes= filetype)

        former_csv_file = pd.read_csv(Basic_inf.csv_file_path,encoding='utf-8-sig')

        return former_csv_file

    #基本的な利用状況の確認
    def judge(df: pd):

        index = ['支払手法']

        #支払方法を確認する
        print('-------------------------\n')
        print('支払方法の様子\n')

        if df['支払方法'].str.contains('1回払い').any():
            print('1回払いがあります\n')
            df_bool_lump_sum_payment = df['支払方法'] == '1回払い'
            count_lump_sum_payment = df_bool_lump_sum_payment.sum()
            print('1回払いの回数は' + str(count_lump_sum_payment) + '回です\n')

        if df['支払方法'].str.contains('リボ払い').any():
            print('リボ払いがあります\n')
            df_bool_revolving_payment = df['支払方法'] == 'リボ払い'
            count_revolving_payment = df_bool_revolving_payment.sum()
            print('リボ払いの回数は' + str(count_revolving_payment) + '回です\n')
        else:
            print('リボ払いはありません\n')

        if df['支払方法'].str.contains('分割払い').any():
            print('分割払いがあります\n')
            df_bool_payment_in_installments = df['支払方法'] == 'リボ払い'
            count_payment_in_installments = df_bool_payment_in_installments.sum()
            print('リボ払いの回数は' + str(count_payment_in_installments) + '回です\n')
        else:
            print('分割払いはありません\n')
            

        #誰がカードを利用したか確認する(基本的に本人か家族かのいずれかが記録されている)
        print('-------------------------\n')
        print('使用者の情報\n')

        if df['利用者'].str.contains('本人').any():
            print('本人の使用履歴があります\n')
            df_bool_myself = df['利用者'] == '本人'
            count_myself = df_bool_myself.sum()
            print('本人の利用回数は' + str(count_myself) + '回です\n')

        else:
            print('本人の使用履歴はありません\n')

        if df['利用者'].str.contains('家族').any():
            print('家族の使用履歴があります\n')
            df_bool_family = df['利用者'] == '家族'
            count_family = df_bool_family.sum()
            print('家族の利用回数は' + str(count_family) + '回です\n')

        else:
            print('家族の使用履歴はありません\n')

        print('-------------------------\n')
        print('支払手数料の確認\n')

        df_additional_payment = df['支払手数料'].sum().any()

        if df_additional_payment > 0:
            print('今月の支払手数料は' + str(df_additional_payment) + '円です\n')
        else:
            print('今月の支払手数料はありません\n')

        print('---------------------------------------')
        print('基本的な利用状況は以上です。\n')
        print('詳細は保存したcsvファイルに記述されています。\n')

    def edit_csv_file(df: pd):

        #店舗別に支払総額をまとめる
        df_payment_per_category = df.groupby('利用店名・商品名')['支払総額'].sum()

        #店舗別支払総額を降順に並べ、indexを再設定
        df_payment_per_category = df_payment_per_category.sort_values(ascending=False).rename_axis('index').reset_index()

        #列名の設定
        df_payment_per_category.columns = ['利用店舗','支払額']

        #店舗別に利用回数を調べる。
        df_count_payment = df.groupby('利用店名・商品名')['利用店名・商品名'].count()

        #indexを変更し、リセット
        df_count_payment = df_count_payment.rename_axis('index').reset_index()

        #列名の設定
        df_count_payment.columns = ['利用店舗','支払回数']

        #列名に基づいてデータを結合
        new_csv_file = pd.merge(df_payment_per_category,df_count_payment ,on = '利用店舗')

        #indexを1から割り当てる。これで店舗別支払額の順位を示せる。
        new_csv_file.index = np.arange(1,len(new_csv_file)+1)

        #店舗別の平均支払金額を算出。小数点以下切り捨て
        new_csv_file['平均支払金額'] = new_csv_file['支払額'] // new_csv_file['支払回数']

        return new_csv_file

    #保存
    def save_file(df: pd):
        title = '保存先を選択してください。'
        filetype = [('xlsxファイル','*.xlsx')]

        #明細のファイル名がenavi202208(????).csv等になっているので、正規表現で年月以降を拾っている
        former_target = 'enavi'
        former_idx = Basic_inf.csv_file_path.find(former_target)

        #明細のファイル名がenavi202208(????).csv等になっているので、正規表現で.csvまで拾っている
        latter_target = '.csv'
        latter_idx = Basic_inf.csv_file_path.find(latter_target)

        #上記で得たindexをスライスで指定し、ファイル名のうち年月日のみを取得し、デフォルトのファイル名としている。
        r = Basic_inf.csv_file_path[former_idx+5:latter_idx]

        initialfile = r
        defaultextension='.xlsx'
        #saved_filename = filedialog.asksaveasfilename(title = title, filetypes = filetype,initialfile = initialfile,defaultextension = defaultextension)

        saved_filename = r + defaultextension

        df.to_excel(saved_filename,index = True,header = True,sheet_name = '店舗別利用状況')

        Basic_inf.new_file_path = saved_filename
        
        
        
    def main():

        #プロンプトにメッセージ表示
        print('クレジットカードの利用明細を分析します。')
        print('エクスプローラーが開いたらダウンロードした利用明細を選択してください。')
        print('エクスプローラーが表示されるまでしばらくお待ちください。')
        
        #ファイル取得
        try:
            former_csv_file = Analyze_credit_card.load_csv_file()
        except FileNotFoundError:
            print('ファイルが見つかりません。')
        
        #指定したファイルが楽天クレジットカード利用明細でない場合、pandasでの処理でエラーが発生する。
        #また、楽天クレジットカード利用明細の仕様が少しでも変更されてもエラーが発生する。
        try:
            Analyze_credit_card.judge(former_csv_file)

            new_csv_file = Analyze_credit_card.edit_csv_file(former_csv_file)

            Analyze_credit_card.save_file(new_csv_file)

            Show_graph.depict_purchase_amount(new_csv_file)

            Show_graph.depict_purchase_frequency(new_csv_file)

            Show_graph.depict_mean_purchase(new_csv_file)

        except Exception:
            print('指定したファイルが楽天クレジットカードの利用明細ではありません。')


        print('分析が終了しました。')

        #終了条件
        while True:
            enter = input('Enterを押したらプロンプトが閉じられます')
            if enter =='':
                break

#実行
if __name__ == '__main__':
    Analyze_credit_card.main()