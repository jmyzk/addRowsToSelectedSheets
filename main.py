######### I created this code on the google Colab platform. #########

# !pip install smartsheet-python-sdk
import smartsheet
import json
access_token = "######################"
smartsheet_client = smartsheet.Smartsheet(access_token)

######################## fix form error for 2021 normal assinments ##########
def fix_this_sheet(shinkoku_sheet_id):
    sheet_to_fix_value = smartsheet_client.Sheets.get_sheet(shinkoku_sheet_id)
 
    # 項目のcolumnIdを取得 Select columnId of an item
    for column in sheet_to_fix_value.columns:
        if column.title ==  "項目":
            koumoku_column_id = column.id
            print(column.title)

    # アップデート用のセル式データの作成  Creating cell formula data for update
    shusei_new_cell = smartsheet.models.Cell()
    shusei_new_cell.column_id = koumoku_column_id
    shusei_new_cell.value = ""

    item_value = ["1","3","4","5"]
    # 2121以降の項目の1,3,4,5を見つける Find 1,3,4,5 for items after 2121
    rows_to_update = []
    for row in sheet_to_fix_value.rows:
        if row.get_column(koumoku_column_id).display_value in item_value and row.created_at.year == 2021:
            new_row = smartsheet.models.Row()
            new_row.id = row.id
            new_row.cells.append( shusei_new_cell)
            rows_to_update.append(new_row)
    updated_row = smartsheet_client.Sheets.update_rows(
    shinkoku_sheet_id,      # sheet_id
    rows_to_update
    )


######################### add formula to rows ############################### 
# 式を追加する行の項目名リスト・タプル
item_name = ("通常業務","課題業務","その他業務",
              "通常業務の業績","課題業務の業績","その他業務の業績",
             "G1の能力・態度評価項目","G23の能力・態度評価項目","G123共通の知識・技能評価項目",
             "G45の能力評価項目","G45の知識・技能評価項目")

def add_formula_to_this_sheet(shinkoku_sheet_id):
    sheet_to_add_formula = smartsheet_client.Sheets.get_sheet(shinkoku_sheet_id)
    # 通常業務、課題業務、その他業務、能力等評価項目 の該当項目のフォーミュラを更新 
    # Updated the formula for the relevant items in the evaluation items such as normal work, task work, other work, ability, etc.
    #   自己、部門
    #       =IFERROR(SUM(CHILDREN(自己評価@row)) / COUNTM(CHILDREN(自己評価@row)), 0)
    #       =IFERROR(SUM(CHILDREN(部門評価@row)) / COUNTM(CHILDREN(部門評価@row)), 0)

    # 項目、自己、部門、全体調整　のcolumnIdを取得  Get columnId of item, self, department, overall adjustment
    for column in sheet_to_add_formula.columns:
        if column.title ==  "項目":
            koumoku_column_id = column.id
        if column.title ==  "自己評価":
            jiko_column_id = column.id
        if column.title ==  "部門評価":
            bumon_column_id = column.id
        # if column.title ==  "全社調整":
        #  zensha_column_id = column.id

    # アップデート用のセル式データの作成 Creating cell formula data for update
    jiko_new_cell = smartsheet.models.Cell()
    jiko_new_cell.column_id = jiko_column_id
    jiko_new_cell.formula = "=IFERROR(SUM(CHILDREN(自己評価@row)) / COUNTM(CHILDREN(自己評価@row)), 0)"
    bumon_new_cell = smartsheet.models.Cell()
    bumon_new_cell.column_id = bumon_column_id
    bumon_new_cell.formula = "=IFERROR(SUM(CHILDREN(部門評価@row)) / COUNTM(CHILDREN(部門評価@row)), 0)"
    
    # 本年度上期の行IDを見つける Find the row ID for the first half of this year
    rows_to_update = []
    for row in sheet_to_add_formula.rows:
        if row.get_column(koumoku_column_id).display_value in item_name and row.created_at.year == 2021:
            new_row = smartsheet.models.Row()
            new_row.id = row.id
            new_row.locked = True
            new_row.cells.append(jiko_new_cell)
            new_row.cells.append(bumon_new_cell)
    #        new_row.cells.append(zensha_new_cell)
            rows_to_update.append(new_row)

    updated_row = smartsheet_client.Sheets.update_rows(
    shinkoku_sheet_id,      # sheet_id
    rows_to_update
    )

#########################2021-2031追加 add ###########################
# origin_sheet_id = to copy rows from                      
origin_sheet_id = 3061110183094148  # 自己申告年度追加用シート
origin_sheet = smartsheet_client.Sheets.get_sheet(origin_sheet_id)

# 項目のcolumnIdを取得 Get columnId of item
for column in origin_sheet.columns:
    if column.title ==  "分類":
        koumoku_column_id = column.id
    
# 本年度上期の行IDを見つける Find the row ID for the first half of this year
origin_row_ids =[]
for row in origin_sheet.rows:
    for cell in row.cells:
        if cell.column_id==koumoku_column_id:
            if cell.value == "年度":
                origin_row_ids.append(row.id)
for id in origin_row_ids:
    print(id)

def copy_rows_from_origin_sheet_to_this_sheet(target_id):
    response = smartsheet_client.Sheets.copy_rows(
    origin_sheet_id,
    smartsheet.models.CopyOrMoveRowDirective({
        'row_ids': origin_row_ids,
        'to': smartsheet.models.CopyOrMoveRowDestination({
        'sheet_id': target_id,     # 対象自己申告シート
        })
    },
    ),
        ['children'] 
    )

############### test ########################################################
sheetId = [477910541854596, 1407738883401604,5911063632865156,1690175932786564]
for target_id in sheetId:
    # copy_rows_from_origin_sheet_to_this_sheet(target_id)
    # add_formula_to_this_sheet(target_id)
    dummy = 'dummy'
print('copy rows from other sheet test done')


######################## all sheet real ###################################
response = smartsheet_client.Sheets.list_sheets(include_all=True)
sheets = response.data

for sheet in sheets:
  data = json.loads(str(sheet))
  name = data['name']
  shinkoku_sheet_id = data['id']
  if '自己申告" のコピー' in name:
    print(name+str(shinkoku_sheet_id)+'  -- ２１年度以降のfomula追加')
    ###### call one of the functions by removing commnet ######
    # copy_rows_from_origin_sheet_to_this_sheet(shinkoku_sheet_id)
    # add_formula_to_this_sheet(shinkoku_sheet_id)
    # fix_this_sheet(shinkoku_sheet_id)

print('fix all self-assesment-sheet test done')

