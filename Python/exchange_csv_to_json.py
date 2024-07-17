import csv
import json

# 入力CSVファイルのパス
input_csv = 'input.csv'

# OpenAPI仕様の初期テンプレート
oas_template = {
    "openapi": "3.0.0",
    "info": {
        "title": "API Document",
        "version": "1.0.0"
    },
    "paths": {},
    "components": {
        "schemas": {}
    }
}

def nest_schema(hierarchy, current_level, properties):
    if not hierarchy:
        return properties

    next_level = hierarchy.pop(0)
    if next_level not in current_level:
        current_level[next_level] = {
            "type": "object",
            "properties": {}
        }

    current_level[next_level]["properties"].update(properties)
    return current_level

# CSVファイルの読み取り
with open(input_csv, mode='r', encoding='utf-8') as file:
    reader = csv.DictReader(file)
    components_schemas = {}
    path_specs = {
        "/examplePath": {
            "get": {
                "summary": "Example summary",
                "responses": {
                    "200": {
                        "description": "A successful response",
                        "content": {
                            "application/json": {
                                "schema": {
                                    "$ref": "#/components/schemas/Response"
                                }
                            }
                        },
                        "headers": {}
                    }
                }
            }
        }
    }
    response_properties = {}
    hierarchy_map = {}

    for row in reader:
        hierarchy = row['階層'].split('.')
        if row['項目（英名）'] == 'Content-Type':
            path_specs["/examplePath"]["get"]["responses"]["200"]["headers"]["Content-Type"] = {
                "schema": {
                    "type": "string",
                    "example": row['値例']
                },
                "description": row['項目説明']
            }
        elif row['項目（英名）'] == 'StatusCode':
            path_specs["/examplePath"]["get"]["responses"]["200"]["headers"]["StatusCode"] = {
                "schema": {
                    "type": "integer",
                    "example": int(row['値例'])
                },
                "description": row['項目説明']
            }
        else:
            nested_property = {
                row['項目（英名）']: {
                    "type": "string" if row['データ型'] == 'string' else "integer",
                    "description": row['項目説明']
                }
            }
            response_properties = nest_schema(hierarchy, response_properties, nested_property)

    components_schemas["Response"] = {
        "type": "object",
        "properties": response_properties
    }

# OpenAPIテンプレートに生成されたコンポーネントとパスを追加
oas_template["components"]["schemas"] = components_schemas
oas_template["paths"] = path_specs

# JSONファイルに出力
with open('openapi_spec.json', 'w', encoding='utf-8') as json_file:
    json.dump(oas_template, json_file, ensure_ascii=False, indent=4)

print("OpenAPI specification has been generated successfully!")