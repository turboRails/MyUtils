import sys
import subprocess
import openpyxl
import re

## ソース出力フォルダー
SRC_OUTPUT_DIRECTORY = '../../'

#{'bigint': 255, 'varchar(n)': 170, 'tinyint': 116, 'integer': 113, 'date': 34 , 'boolean': 57, 'timestamp': 120, 'time': 5, 'Integer': 2, 'datetime': 1}

#DB設計書の型：MySqlの型
mysqlTypeConv = {'tinyint':'TINYINT', 'bigint':'BIGINT', 'integer':'INTEGER', 'Integer':'INTEGER', 'varchar(n)':'STRING', 'boolean':'BOOLEAN', 'timestamp':'DATE', 'date':'DATE', 'time':'TIME', 'datetime': 'DATE'}
#DB設計書の型：node.jsの型
nodeTypeConv = {'tinyint':'number', 'bigint':'number', 'integer':'number', 'Integer':'number', 'varchar(n)':'string', 'boolean':'boolean', 'timestamp':'Date', 'date':'Date', 'time':'Date', 'datetime': 'Date'}

def snake_to_lower_camel(text):
    return re.sub('_([a-zA-Z0-9])', lambda m: m.group(1).upper(), text)

def snake_to_upper_camel(text):
    return re.sub('_([a-zA-Z0-9])', lambda m: m.group(1).upper(), text.capitalize())

def top_lower(text):
    return text[0].lower() + text[1:]

def top_upper(text):
    return text[0].upper() + text[1:]

################
# メイン処理
################

## コマンドパラメータ取得
args = sys.argv
if 2 != len(args):
    print('invalid args')
    print('usage:')
    print("> python " + args[0] + " path/to/DB設計.xlsx")
    sys.exit("Error")

## ワークブックOPEN
wb = openpyxl.load_workbook(args[1])
ws = wb.sheetnames

tsvfile = open('db.tsv', 'w')

tables = {}
uniqueKey = {}
fKey = {}
hasMany = {}
belongsTo = {}

#自動生成できないリレーションを追加
hasMany['control_item_api'] = []
hasMany['control_item_api'].append('\n'.join([
    "    models.controlItemApi.hasMany(models.roleControlItem, {",
    "      onUpdate: 'CASCADE',",
    "      foreignKey: 'controlItemId',",
    "      targetKey: 'id'",
    "    });\n"
]))
hasMany['role_control_item'] = []
hasMany['role_control_item'].append('\n'.join([
    "    models.roleControlItem.hasMany(models.controlItemApi, {",
    "      onUpdate: 'CASCADE',",
    "      foreignKey: 'controlItemId',",
    "      targetKey: 'id'",
    "    });\n"
]))

#作成しないシート名
skipSheets = [
'SKU変換',
'SKU変換明細',
'トラック着車予約',
'トラック着車予約と入荷予定明細',
'CSVアップロード履歴詳細'
]

tt = {}
## ワークシート順次処理
for shname in wb.sheetnames:
    sheet = wb[shname]
    if 'テーブル名(論理)' in str(sheet.cell(row=4, column=1).value) and (shname not in skipSheets):
        #print(sheet.cell(row=4, column=11).value)
        tableName = sheet.cell(row=5, column=11).value
        tables[tableName] = []
        uniqueKey[tableName] = []
        uniqueKeyFlg = False
        fKey[tableName] = []
        fKeyFlg = False
        colFlg = True
        if tableName not in belongsTo:
            belongsTo[tableName] = []

        maxRow = wb[shname].max_row + 1
        for i in range(10, maxRow):

            if sheet.cell(row=i, column=1).value == '複合一意キー':
                uniqueKeyFlg = True
                continue
            if sheet.cell(row=i, column=1).value == '外部キー':
                fKeyFlg = True
                continue

            if not sheet.cell(row=i, column=1).value:
                uniqueKeyFlg = False
                colFlg = False
                fKeyFlg = False
                continue

            if uniqueKeyFlg and sheet.cell(row=i, column=1).value != 'なし':
                uniqueKey[tableName].append(sheet.cell(row=i, column=1).value)
                continue

            if fKeyFlg:
                if sheet.cell(row=i, column=1).value != 'カラム名 (論理名)' and sheet.cell(row=i, column=1).value != 'なし':
                    foreignKey = sheet.cell(row=i, column=11).value
                    refTable = sheet.cell(row=i, column=21).value
                    sourceKey = sheet.cell(row=i, column=31).value
                    fKey[tableName].append([
                                      foreignKey #0
                                    , refTable #1
                                    , sourceKey #2
                                    ])
                    #print(refTable)
                    foreignKey = snake_to_lower_camel(foreignKey)
                    sourceKey = snake_to_lower_camel(sourceKey)
                    if refTable not in hasMany:
                        hasMany[refTable] = []
                    hasMany[refTable].append('\n'.join([
                        "    models." + snake_to_lower_camel(refTable) + ".hasMany(models." + snake_to_lower_camel(tableName) + ", {",
                        "      onUpdate: 'CASCADE',",
                        "      foreignKey: '" + (foreignKey) + "',",
                        "      sourceKey: '" + (sourceKey) + "'",
                        "    });\n"
                    ]))
                    belongsTo[tableName].append('\n'.join([
                        "    models." + snake_to_lower_camel(tableName) + ".belongsTo(models." + snake_to_lower_camel(refTable) + ", {",
                        "      onUpdate: 'CASCADE',",
                        "      foreignKey: '" + (foreignKey) + "',",
                        "      targetKey: '" + (sourceKey) + "'",
                        "    });\n"
                    ]))
                continue

            if not colFlg:
                continue

            colmunName = snake_to_lower_camel(sheet.cell(row=i, column=11).value)
            field = sheet.cell(row=i, column=11).value
            colmunType = sheet.cell(row=i, column=21).value
            colmunLength = sheet.cell(row=i, column=27).value
            if not colmunLength:
                colmunLength = sheet.cell(row=i, column=30).value
            primaryKey = sheet.cell(row=i, column=35).value
            notNull = sheet.cell(row=i, column=38).value
            autoIncrement = sheet.cell(row=i, column=53).value
            defaultValue = sheet.cell(row=i, column=59).value
            name = sheet.cell(row=i, column=1).value
            tables[tableName].append([
                                      colmunName #0
                                    , colmunType #1
                                    , primaryKey #2
                                    , autoIncrement #3
                                    , notNull #4
                                    , field #5
                                    , defaultValue #6
                                    , name #7
                                    , colmunLength #8
                                    ])
            if colmunType not in tt:
                tt[colmunType] = 0
            tt[colmunType] += 1

            tsvfile.write(sheet.cell(row=4, column=11).value)
            tsvfile.write(".")
            tsvfile.write(sheet.cell(row=i, column=1).value)
            tsvfile.write("\t")
            tsvfile.write(sheet.cell(row=5, column=11).value)
            tsvfile.write(".")
            tsvfile.write(sheet.cell(row=i, column=11).value)
            tsvfile.write("\t")
            tsvfile.write(snake_to_upper_camel(sheet.cell(row=5, column=11).value))
            tsvfile.write(".")
            tsvfile.write(snake_to_lower_camel(sheet.cell(row=i, column=11).value))
            tsvfile.write("\n")
tsvfile.close()

dbInterfaceFile = open(SRC_OUTPUT_DIRECTORY + 'sb-scm-api/src/share/mysqlModels/dbInterface.ts', 'w')
importList = ['import { Sequelize } from \'sequelize\';\n']
interfaceList = ['  sequelize: Sequelize;\n']
indexList = ['import { sequelize } from \'../modules/dbConfig\';\n']
exportList = ['  sequelize: sequelize']
for tableName in tables.keys():
    importList.append('import { ' + snake_to_upper_camel(tableName) + 'Model } from \'./tables/' + snake_to_lower_camel(tableName) + '\';\n')
    interfaceList.append('  ' + snake_to_lower_camel(tableName) + ': ' + snake_to_upper_camel(tableName) + 'Model;\n')
    indexList.append('import { create' + snake_to_upper_camel(tableName) + 'Model } from \'./tables/' + snake_to_lower_camel(tableName) + '\';\n')
    exportList.append('  ' + snake_to_lower_camel(tableName) + ': create' + snake_to_upper_camel(tableName) + 'Model(sequelize)')

importList.sort()
interfaceList.sort()
dbInterfaceFile.write(''.join(importList))
dbInterfaceFile.write('\n')
dbInterfaceFile.write('export interface DbInterface {\n')
dbInterfaceFile.write(''.join(interfaceList))
dbInterfaceFile.write('}\n')
dbInterfaceFile.close()
#####
indexList.sort()
exportList.sort()
indexFile = open(SRC_OUTPUT_DIRECTORY + 'sb-scm-api/src/share/mysqlModels/index.ts', 'w')
indexFile.write(''.join(indexList) + '\n')
indexFile.write('\n')
indexFile.write('export const db = {\n')
indexFile.write(',\n'.join(exportList) + '\n')
indexFile.write('};\n')
indexFile.close()
#####
for tableName, v in tables.items():

    file = open(SRC_OUTPUT_DIRECTORY + 'sb-scm-api/src/share/mysqlModels/tables/' + snake_to_lower_camel(tableName) + '.ts', 'w')
    daoFile = open(SRC_OUTPUT_DIRECTORY + 'sb-scm-api/src/share/dao/' + snake_to_lower_camel(tableName) + '.ts', 'w')

    typeList = ['Model', 'Sequelize']
    for col in v:
        dataType = mysqlTypeConv.setdefault(col[1], 'TODO')
        if dataType not in typeList:
            typeList.append(dataType)
    typeList.sort()

    file.write('import {\n')
    file.write('  ' + ',\n  '.join(typeList) + '\n')
    file.write('} from \'sequelize\';\n')
    file.write('\n')

    file.write('export interface ' + snake_to_upper_camel(tableName) + 'Attributes {\n')
    for col in v:
        file.write('  /** ' + col[7] + ' */\n')
        file.write('  ' + col[0] + '?: ' + nodeTypeConv.setdefault(col[1], 'TODO') + ';\n')

    file.write('}\n')
    file.write('\n')

    file.write('export class ' + snake_to_upper_camel(tableName) + 'Model extends Model {\n')
    if (tableName in hasMany and len(hasMany[tableName]) > 0) or (tableName in belongsTo and len(belongsTo[tableName]) > 0):
        file.write('  static associate = models => {\n')
        if tableName in hasMany:
            file.write('\n'.join(hasMany[tableName]))
            file.write('\n')
        if tableName in belongsTo:
            file.write('\n'.join(belongsTo[tableName]))
            file.write('\n')
        file.write('  };\n')
    file.write('}\n')
    file.write('\n')
    file.write('export const create' + snake_to_upper_camel(tableName) + 'Model = (sequelize: Sequelize): any => {\n')
    file.write('  ' + snake_to_upper_camel(tableName) + 'Model.init(\n')
    file.write('    {\n')
    column = []
    for col in v:
        text = []
        text.append('      ' + col[0] + ': {\n')
        text.append('        type: ' + mysqlTypeConv.setdefault(col[1], 'TODO') + ',\n')
        if col[2]:
            text.append('        primaryKey: ' + ('true' if col[2] else 'false') + ',\n')
        if col[3]:
            text.append('        autoIncrement: ' + ('true' if col[3] else 'false') + ',\n')
        text.append('        allowNull: ' + ('false' if col[4] else 'true') + ',\n')
        if col[6]:
            text.append('        defaultValue: \'' + col[6] + '\',\n')
        text.append('        field: \'' + col[5] + '\'\n')
        text.append('      }')
        column.append("".join(text))

    file.write(',\n\n'.join(column) + '\n')
    file.write('    },\n')
    file.write('    {\n')
    file.write('      sequelize,\n')
    file.write('      freezeTableName: true,\n')
    file.write('      timestamps: false,\n')
    file.write('      tableName: \'' + tableName + '\',\n')
    file.write('      modelName: \''+ snake_to_lower_camel(tableName) + '\'\n')
    file.write('    }\n')

    file.write('  );\n')
    file.write('  return ' + snake_to_upper_camel(tableName) + 'Model;\n')
    file.write('};\n')
    file.close()

    daoFile.write('import { db } from \'../mysqlModels\';\n')
    daoFile.write('import { Transaction } from \'sequelize\';\n')
    daoFile.write('import { ' + snake_to_upper_camel(tableName) + 'Attributes } from \'../mysqlModels/tables/' + snake_to_lower_camel(tableName) + '\';\n')
    daoFile.write('import { sequelizeUtil } from \'../modules/sequelizeUtil\';\n')
    daoFile.write('\n')
    daoFile.write('export class ' + snake_to_upper_camel(tableName) + 'Dao {\n')
    daoFile.write('  static findAll(where' + snake_to_upper_camel(tableName) + ': any, tx?: Transaction) {\n')
    daoFile.write('    return db.' + snake_to_lower_camel(tableName) + '.findAll({\n')
    daoFile.write('      include: [].concat(sequelizeUtil.includeCommonColumn(\'' + snake_to_lower_camel(tableName) + '\')),\n')
    daoFile.write('      where: where' + snake_to_upper_camel(tableName) + ',\n')
    daoFile.write('      transaction: tx\n')
    daoFile.write('    });\n')
    daoFile.write('  }\n')
    daoFile.write('\n')
    daoFile.write('  static insert(attribute: ' + snake_to_upper_camel(tableName) + 'Attributes, tx?: Transaction) {\n')
    daoFile.write('    return db.' + snake_to_lower_camel(tableName) + '.create(attribute, { transaction: tx });\n')
    daoFile.write('  }\n')
    daoFile.write('\n')
    daoFile.write('  static update(attribute: ' + snake_to_upper_camel(tableName) + 'Attributes, tx?: Transaction) {\n')
    daoFile.write('    return db.' + snake_to_lower_camel(tableName) + '.update(attribute, {\n')
    daoFile.write('      where: {\n')
    daoFile.write('        id: attribute.id\n')
    daoFile.write('      },\n')
    daoFile.write('      transaction: tx\n')
    daoFile.write('    });\n')
    daoFile.write('  }\n')
    daoFile.write('}\n')
    daoFile.close()

#スタイルチェック+コード整形
args = ['npm', 'run', 'lint:fix']
try:
    res = subprocess.check_call(args)
    print(res)
except:
    print("Error.")

print('end!')
