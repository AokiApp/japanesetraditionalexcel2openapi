import pandas as pd
import os
import yaml

excel_file = '/home/user/codes/japanesetraditionalexcel2openapi/Sample/Sample.xlsx'
#df_sheet_multi = pd.read_excel(excel_file, sheet_name=['入力項目定義','出力項目定義'], header=6)
read_sheet_header = 7
df_sheet_input_fields = pd.read_excel(excel_file, sheet_name='入力項目定義', header=read_sheet_header)
df_sheet_output_fields = pd.read_excel(excel_file, sheet_name='出力項目定義', header=read_sheet_header)

# 読み込んだシートに対して処理を行う
# Excel方眼紙であるために発生するUnnamedなヘッダーのカラムを削除し、冒頭2行に発生しがちな値がすべてNaNの行を取り除く
# そもそも項目名で取り出すしいらない処理かもしれない
## 少なくともNaNを削除する.dropna(how='all')はあったほうがスムーズに値を取り出せそう
df_sheet_input_fields = df_sheet_input_fields.loc[:, ~df_sheet_input_fields.columns.str.contains('^Unnamed')].dropna(how='all')
df_sheet_output_fields = df_sheet_output_fields.loc[:, ~df_sheet_output_fields.columns.str.contains('^Unnamed')].dropna(how='all')

# sumary用にファイル名から単語を取得
excel_file_name = os.path.basename(excel_file)
## 拡張子を除いたファイル名を取得
summary_text = os.path.splitext(excel_file_name)[0]
# アンダーバー以前の部分を取得
if '_' in summary_text:
    summary_text = summary_text.split('_')[0]
else:
    summary_text = summary_text

# Excelから取得できない部分の設定
http_request_set = "post" #API設計書から判断する必要がある
endpoint_path_set = "example_endpoint"
operation_id_set = "example_operationId"
response_code = "200"#Excelファイルから取得した値から指定するか、なければ200指定？というか記載が無いだけで400とか他も設定必要では->200決め打ちでいいっぽい
responses_schema_type = "object"

# Excelファイル内からレスポンスヘッダのコンテンツタイプ取得													
input_fields_content_type = df_sheet_input_fields[df_sheet_input_fields['項目名（英語）'].str.contains('content-type', case=False)]['値の例'].iloc[0]
output_fields_content_type = df_sheet_output_fields[df_sheet_output_fields['項目名（英語）'].str.contains('content-type', case=False)]['値の例'].iloc[0]

#実行結果の確認。随時削除予定
# print(summary_text)
# print(input_fields_content_type)
# print("入力項目定義")
# print(df_sheet_input_fields)
# print("出力項目定義")
# print(df_sheet_output_fields)

# ここから必要な項目についての話というか作りたい最終的なフォーマットの話
"""
Excelに含まれるデータは
No,階層,項目名（日本語）,項目名（英語）,データ型,桁数,必須,入出力区分,種別,項目の説明,値の例
1.GETかPOSTかの設定は変数を置いて各自で行う
2.パスの設定も各自で行う
3. Excelから取得可能な情報は使って埋める
以下は想定しているyamlファイルの完成形。プログラムで入れる部分は${}で示す。
path:
  ${endpoint_path_set}
    ${http_request_set}:
        tags:
        - 多分ファイル名で申請アプリ_XXX_でグループ化されているならあったほうがいいもの。いったん無視でいいかも
        summary:${summary_text}
        description: 保留
        operationId: ${operation_id_set}
        ---メソッドがpost(get以外)の場合---
        requestBody:
            description:未定
            requiered: 
                - [必須項目がyになっている項目名を列挙]
            content:
                ${input_fields_content_type}
                schema:
                    type: object
                    properties:
                        [項目名（英名）]:
                            type: [データ型]
                            description: [項目説明]
                            example: [値例]
        ---
        ---メソッドがgetでクエリパラメータやパスパラメータを持つ場合---
        parameters:
            - name: "クエリパラメータ名"
              in: "query"
              required: true/false
              schema:
                type: 
                description: 
            - name: "パスパラメータ名"
            ....
        ---
        response:
            200:
                description:
                    'どうやって埋めるか保留'
                content:
                    [値例]
                    schema:
                    type: object
                    properties:
                        英語項目名1:
                            type: [データ型]
                            description: [項目説明]
                            example: [値例]
                            ---もしネスト（次の行の階層が2以上なら）---
                            properties:
                                英語項目名:
                                    type: [データ型]
                                    description: [項目説明]
                                    example: [値例]

                        英語項目名2:
                        ...
                    required:
                        必須項目がyになっている項目名を列挙
                        必須項目がyになっている項目名を列挙
    
"""
# ここから実装
# OASの基本構造
oas = {
    "openapi": "3.0.0",
    "info": {
        "title": "API Documentation",
        "version": "1.0.0"
    },
    "paths": {
        endpoint_path_set: {
            http_request_set: {
                "summary": summary_text,  # 例: 保留
                "description": "どう取得するか保留",
                "operationId": operation_id_set,  # 例: 各自設定
                "responses": {
                    response_code: {
                        "description": "レスポンスの説明をどう取得するか保留",
                        "content": {
                            output_fields_content_type: {
                                "schema": {
                                    "type": responses_schema_type,
                                    "properties": {}
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

# POSTメソッドの場合、requestBodyを追加
if http_request_set == "post":
    oas["paths"][endpoint_path_set][http_request_set]["requestBody"] = {
        "description": "未定",
        "required": [],
        "content": {
            input_fields_content_type: {
                "schema": {
                    "type": "object",
                    "properties": {}
                }
            }
        }
    }   

# ネスト用のスタック
#stack = []

# リクエストボディ用の更新
for index, row in df_sheet_input_fields.iterrows():
    # ヘッダー行はスキップ
    if row["種別"].lower() == "header":
        continue
    field_name_english = row["項目名（英語）"]
    data_type = row["データ型"]
    description = row["項目の説明"]
    example = row["値の例"]
    required = str(row["必須"]).lower() == "y"
    hierarchy = row["階層"]


    prop = {
        "type": data_type,
        "description": description
    }
    # example が NaN でない場合のみ追加
    if not pd.isna(example):
        prop["example"] = example

    oas["paths"][endpoint_path_set][http_request_set]["requestBody"]["content"][input_fields_content_type]["schema"]["properties"][field_name_english] = prop

    if required:
        oas["paths"][endpoint_path_set][http_request_set]["requestBody"].setdefault("required", []).append(field_name_english)

# responsesの更新
# ネスト用のスタック
stack = []
current_properties = oas["paths"][endpoint_path_set][http_request_set]["responses"][response_code]["content"][output_fields_content_type]["schema"]["properties"]
# レスポンスボディ用の更新
for index, row in df_sheet_output_fields.iterrows():
    # ヘッダー行はスキップ
    if row["種別"].lower() == "header":
        continue
    field_name_english = row["項目名（英語）"]
    data_type = row["データ型"]
    description = row["項目の説明"]
    example = row["値の例"]
    hierarchy = int(row["階層"])
    # 空白の場合.lowerがエラーになるので文字列に変換してから判定
    required = str(row["必須"]).lower() == "y"

    prop = {
        "type": data_type,
        "description": description
    }

    # プロパティの設定。listとなっていたらarrayに変換
    if data_type == "list":
        prop = {
            "type": "array",
            "description": description,
            "items": {
                "type": "object",
                "properties": {}
            }
        }
    else:
        prop = {
            "type": data_type,
            "description": description
        }

    # example は NaN でない場合のみ追加
    if not pd.isna(example):
        prop["example"] = example
    
    # ネストの管理: 現在の階層より高い階層の要素をスタックから取り除く
    while stack and stack[-1]['hierarchy'] >= hierarchy:
        stack.pop()

    # 現在の階層のプロパティに追加
    if stack:
        parent_properties = stack[-1]['object']['properties']
        parent_properties[field_name_english] = prop
    else:
        current_properties[field_name_english] = prop

    # properties の追加は object 型要素にのみ適用
    if data_type == "object":
        prop['properties'] = {}
        stack.append({'hierarchy': hierarchy, 'object': prop})
    # #　現状一番うまくいっているがネスト時に不要なpropertiesが含まれてしまうやつ    
    # # 現在の階層のプロパティに追加
    # if stack:
    #     current_properties = stack[-1]['object']['properties']
    # else:
    #     current_properties = oas["paths"][endpoint_path_set][http_request_set]["responses"][response_code]["content"][output_fields_content_type]
    # current_properties[field_name_english] = prop
    # if 'properties' not in prop:
    #     prop['properties'] = {}

    # stack.append({'hierarchy': hierarchy, 'object': prop})
    
    # 必須項目の設定
    if required:
        oas["paths"][endpoint_path_set][http_request_set]["responses"]["200"].setdefault("required", []).append(field_name_english)

output_file = f"{excel_file_name}.yaml"
with open(output_file, "w", encoding='utf8') as yaml_file:
    yaml.dump(oas, yaml_file, allow_unicode=True, sort_keys=False)

print("YAMLファイルに変換が完了しました。")