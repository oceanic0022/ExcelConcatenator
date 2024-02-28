import pandas as pd
import os




#指定した列を読み込む　または　指定した列と行を指定して読み込む
def load_data(file_path, column_name):
    # 特定の行をスキップする条件関数
    skip_func = lambda x: x > 0 and x < 91670
    # Excelファイルを読み込む
    df = pd.read_excel(file_path)
    #df = pd.read_excel(file_path, skiprows=skip_func)
    
    # 指定された列が文字列型でない場合は文字列型に変換
    if df[column_name].dtype != 'object':
        df[column_name] = df[column_name].astype('str')
    
    # 指定された列のデータが空（空文字やNaN）の行を削除
    df = df.dropna(subset=[column_name])  # NaNを削除
    df = df[df[column_name].str.strip().astype(bool)]  # 空文字を削除

    return df[column_name].tolist(), df





#DataFrameをExcelファイルに保存する関数
def save_to_excel(df, output_path):
    directory = os.path.dirname(output_path)
    if not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)
    df.to_excel(output_path, index=False)
    print(f"Results saved to {output_path}")






#ディレクトリ内にあるすべてのエクセルファイルの行を連結させて一つのエクセルファイルにまとめるコード
def concatenate_excel_files(directory, output_filename):
    # ディレクトリ内のすべてのExcelファイルのリストを作成
    excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    
    # 連結するDataFrameのリストを初期化
    dfs = []
    
    # 各Excelファイルを読み込み、リストに追加
    for file in excel_files:
        df = pd.read_excel(os.path.join(directory, file))
        dfs.append(df)
    
    # すべてのDataFrameを一つに連結
    concatenated_df = pd.concat(dfs, ignore_index=True)
    
    # 結果を新しいExcelファイルとして保存
    concatenated_df.to_excel(output_filename, index=False)
