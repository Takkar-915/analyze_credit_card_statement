from operator import index
import pandas as pd
import numpy as np
import openpyxl
from tkinter import filedialog
import matplotlib.pyplot as plt
import io
import os
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
    def analyze_payment_method(df: pd):

        index = ['1回払い','分割払い','リボ払い','本人の使用','家族の使用','手数料']
        about_payment_method = pd.DataFrame(index = index,columns=['利用状況'])
        #about_payment_method.fillna('0',inplace=True)

        #支払方法の確認
        if df['支払方法'].str.contains('1回払い').any():
            df_bool_lump_sum_payment = df['支払方法'] == '1回払い'
            count_lump_sum_payment = df_bool_lump_sum_payment.sum()
            about_payment_method.at['1回払い','利用状況'] = str(count_lump_sum_payment) + '回'

        if df['支払方法'].str.contains('分割払い').any():
            df_bool_payment_in_installments = df['支払方法'] == '分割払い'
            count_payment_in_installments = df_bool_payment_in_installments.sum()
            about_payment_method.at['分割払い','利用状況'] = str(count_payment_in_installments) + '回'
        else:
            about_payment_method.at['分割払い','利用状況'] = '0回'

        if df['支払方法'].str.contains('リボ払い').any():
            df_bool_revolving_payment = df['支払方法'] == 'リボ払い'
            count_revolving_payment = df_bool_revolving_payment.sum()
            about_payment_method.at['リボ払い','利用状況'] = str(count_revolving_payment) + '回'
        else:
            about_payment_method.at['リボ払い','利用状況'] = '0回'

        #誰がカードを利用したか確認する('本人','家族')

        if df['利用者'].str.contains('本人').any():
            df_bool_myself = df['利用者'] == '本人'
            count_myself = df_bool_myself.sum()
            about_payment_method.at['本人の使用','利用状況'] = str(count_myself) + '回'
        else:
            about_payment_method.at['本人の使用','利用状況'] = '0回'

        if df['利用者'].str.contains('家族').any():
            df_bool_family = df['利用者'] == '家族'
            count_family = df_bool_family.sum()
            about_payment_method.at['家族の使用','利用状況'] = str(count_family) + '回'
        else:
            about_payment_method.at['家族の使用','利用状況'] = '0回'

        #支払手数料の確認
        df_additional_payment = df['支払手数料'].sum().any()

        if df_additional_payment > 0:
            about_payment_method.at['手数料','利用状況'] = str(df_additional_payment) + '円'
        else:
            about_payment_method.at['手数料','利用状況'] = '0円'
        
        return about_payment_method

    def analyze_payment_status(df: pd):

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
        about_payment_status = pd.merge(df_payment_per_category,df_count_payment ,on = '利用店舗')

        #indexを1から割り当てる。これで店舗別支払額の順位を示せる。
        about_payment_status.index = np.arange(1,len(about_payment_status)+1)

        #店舗別の平均支払金額を算出。小数点以下切り捨て
        about_payment_status['平均支払金額'] = about_payment_status['支払額'] // about_payment_status['支払回数']

        return about_payment_status

    #保存
    def save_file(about_payment_method,about_payment_status: pd):
    
        #明細のファイル名がenavi202208(????).csv等になっているので、正規表現で年月以降を拾っている
        former_target = 'enavi'
        former_idx = Basic_inf.csv_file_path.find(former_target)

        #明細のファイル名がenavi202208(????).csv等になっているので、正規表現で.csvまで拾っている
        latter_target = '.csv'
        latter_idx = Basic_inf.csv_file_path.find(latter_target)

        #上記で得たindexをスライスで指定し、ファイル名のうち年月日のみを取得し、デフォルトのファイル名としている。
        filename = Basic_inf.csv_file_path[former_idx+5:latter_idx]

        #カレントディレクトリの取得
        current_dir = os.getcwd()
        #カレントディレクトリに移動
        os.chdir(current_dir)

        #出力結果を保存するフォルダ、なければ作る
        new_path = '分析結果'
        if not os.path.exists(new_path):
            os.mkdir(new_path)
        
        #保存先のパスを指定
        basepath = os.path.join(current_dir,new_path)

        #ファイル名の指定
        initialfile = filename
        defaultextension='.xlsx'

        saved_filename = initialfile + defaultextension

        #保存先のパス+保存ファイル名
        saved_path = os.path.join(basepath,saved_filename)

        about_payment_method.to_excel(saved_path,index = True,header = True,sheet_name = '支払方法')

        with pd.ExcelWriter(saved_path,engine='openpyxl',mode ='a') as writer:

            about_payment_status.to_excel(writer,index = True,sheet_name = '店舗別利用状況')
        

        Basic_inf.new_file_path = saved_path
        
        
    def main():

        #プロンプトにメッセージ表示
        print('クレジットカードの利用明細を分析します。')
        print('エクスプローラーが開いたらダウンロードした利用明細を選択してください。')
        print('エクスプローラーが表示されるまでしばらくお待ちください。\n')
        
        #ファイル取得
        try:
            former_csv_file = Analyze_credit_card.load_csv_file()
        except FileNotFoundError:
            print('ファイルが見つかりません。')
        
        #指定したファイルが楽天クレジットカード利用明細でない場合、pandasでの処理でエラーが発生する。
        #また、楽天クレジットカード利用明細の仕様が少しでも変更されてもエラーが発生する。
        try:
            #支払方法等の確認
            about_payment_method= Analyze_credit_card.analyze_payment_method(former_csv_file)

            #店舗別利用状況の確認
            about_payment_status = Analyze_credit_card.analyze_payment_status(former_csv_file)

            #分析内容をexcel形式で保存
            Analyze_credit_card.save_file(about_payment_method,about_payment_status)

            #店舗別の支払総額をグラフで可視化
            Show_graph.depict_purchase_amount(about_payment_status)

            #店舗別の利用回数をグラフで可視化
            Show_graph.depict_purchase_frequency(about_payment_status)

            #店舗別の平均支払金額をグラフを可視化
            Show_graph.depict_mean_purchase(about_payment_status)

            print('----------------------------------------------------------------------------\n')
            print('分析が終了しました。')
            print('分析結果はexeファイルと同じ階層にある「分析結果」というフォルダに保存されています。')

        except Exception:
            print('想定外のエラーが発生しました。指定したファイルが楽天クレジットカードの利用明細(csv形式)かもう一度確認してください。')

        #終了条件
        while True:
            enter = input('Enterを押したらプロンプトが閉じられます')
            if enter =='':
                break

#実行
if __name__ == '__main__':
    Analyze_credit_card.main()