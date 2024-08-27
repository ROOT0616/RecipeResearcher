import discord
from discord import app_commands
from discord.ext import commands
import pandas as pd
import math
import json
import logging
from difflib import get_close_matches
import threading

# ログ設定
logging.basicConfig(filename='bot.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# グローバル変数
config = None
EXCEL_FILE_PATH = None
SPECIAL_ITEMS = None

def load_config():
    global config, EXCEL_FILE_PATH, SPECIAL_ITEMS
    try:
        with open('config.json', 'r', encoding='utf-8') as config_file:
            config = json.load(config_file)
            EXCEL_FILE_PATH = config['EXCEL_FILE_PATH']
            SPECIAL_ITEMS = config['SPECIAL_ITEMS']
            logging.info("設定ファイルが再読み込みされました。")
    except Exception as e:
        logging.error(f"設定ファイルの再読み込み中にエラーが発生しました: {e}")

def save_config(new_config):
    global config
    try:
        with open('config.json', 'w', encoding='utf-8') as config_file:
            json.dump(new_config, config_file, ensure_ascii=False, indent=4)
            config = new_config
            logging.info("設定ファイルが更新されました。")
    except Exception as e:
        logging.error(f"設定ファイルの更新中にエラーが発生しました: {e}")

def load_crafting_data():
    try:
        # Excelファイルのすべてのシートを読み込み
        sheet_dict = pd.read_excel(EXCEL_FILE_PATH, sheet_name=None)
        # シートごとのデータフレームを1つのデータフレームにまとめる
        df = pd.concat(sheet_dict.values(), ignore_index=True)
        return df
    except Exception as e:
        logging.error(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
        return pd.DataFrame()

def calculate_materials(materials, df):
    total_materials = {}
    intermediate_materials = {}

    for item, quantity in materials.items():
        item_data = df[df['完成品名'] == item]

        if item_data.empty:
            continue

        for _, row in item_data.iterrows():
            completion_quantity = row['完成個数']

            for col in df.columns:
                if '材料' in col:
                    material = row[col]

                    # 文字列である場合のみ strip() を実行
                    if isinstance(material, str):
                        material = material.strip()
                    elif pd.notna(material):  # material が NaN でない場合は数値として扱う
                        material = str(material)
                    else:
                        continue

                    required_quantity = row[col.replace('材料', '必要数')]
                    recipe_count = math.ceil(quantity / completion_quantity)
                    total_needed = recipe_count * required_quantity

                    if material in df['完成品名'].values:  # 中間生成物の場合
                        if material not in intermediate_materials:
                            intermediate_materials[material] = 0
                        intermediate_materials[material] += total_needed
                    else:  # 最終材料の場合
                        if material not in total_materials:
                            total_materials[material] = 0
                        total_materials[material] += total_needed

    # 辞書のコピーを作成して、それをループする
    intermediate_materials_copy = intermediate_materials.copy()
    
    for intermediate_item, intermediate_qty in intermediate_materials_copy.items():
        nested_materials, nested_intermediate_materials = calculate_materials({intermediate_item: intermediate_qty}, df)
        for nested_material, nested_qty in nested_materials.items():
            if nested_material not in total_materials:
                total_materials[nested_material] = 0
            total_materials[nested_material] += nested_qty
        for nested_intermediate, nested_intermediate_qty in nested_intermediate_materials.items():
            if nested_intermediate not in intermediate_materials:
                intermediate_materials[nested_intermediate] = 0
            intermediate_materials[nested_intermediate] += nested_intermediate_qty

    return total_materials, intermediate_materials

# Botの設定
intents = discord.Intents.default()
intents.message_content = True
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    logging.info(f'ログインしました: {bot.user.name} ({bot.user.id})')
    await bot.tree.sync()

@bot.tree.command(name='materials', description='Calculate the total amount of materials needed')
@app_commands.describe(materials='List of materials and their quantities in the format "Material:Quantity,Material:Quantity" (e.g., "剛力の宝薬G2:9,魔匠の薬液:3")')
async def materials(interaction: discord.Interaction, materials: str):
    materials_dict = {}
    try:
        # ユーザー入力を辞書に変換
        for arg in materials.split(','):
            item, qty = arg.split(':')
            materials_dict[item.strip()] = int(qty.strip())
    except Exception as e:
        await interaction.response.send_message("エラー: コマンドの形式が正しくありません。例: /materials 剛力の宝薬G2:9,魔匠の薬液:3")
        logging.error(f"材料コマンドの形式エラー: {e}")
        return

    # データフレームを読み込む
    df = load_crafting_data()
    if df.empty:
        await interaction.response.send_message("エラー: Excelファイルを読み込む際に問題が発生しました。")
        return

    # ユーザーが入力した材料が存在しない場合に候補を提案
    unknown_items = [item for item in materials_dict if item not in df['完成品名'].values]
    if unknown_items:
        suggestions = {}
        for item in unknown_items:
            close_matches = get_close_matches(item, df['完成品名'].values, n=5, cutoff=0.5)
            if close_matches:
                suggestions[item] = close_matches

        if suggestions:
            suggestion_message = "以下の材料が見つかりませんでした。近い名前の候補を提示します:\n\n"
            for item, matches in suggestions.items():
                suggestion_message += f"**{item}** :  {', '.join(matches)}\n"
            await interaction.response.send_message(suggestion_message)
            return

    # 材料の計算
    total_materials, intermediate_materials = calculate_materials(materials_dict, df)

    if not total_materials:
        await interaction.response.send_message("エラー: 指定された材料に基づくデータが見つかりませんでした。")
        return

    # 特定のアイテムをリストの最後に持ってくるための処理
    other_materials = {material: qty for material, qty in total_materials.items() if material not in SPECIAL_ITEMS}
    special_materials = {material: qty for material, qty in total_materials.items() if material in SPECIAL_ITEMS}

    # 入力された材料のリストをフォーマット
    materials_list = ""
    for material, quantity in materials_dict.items():
        materials_list += f"{material} を {quantity}個\n"

    # 結果のフォーマット
    result = f"{materials_list}の必要な材料の総量:\n\n"

    # 中間生成物の表示
    intermediate_message = "\n中間生成物:\n"
    for material, qty in intermediate_materials.items():
        intermediate_message += f"{material}  x  {int(qty)}\n"

    # 中間生成物のメッセージ追加
    result += intermediate_message

    # 通常の材料の表示
    result += "\nその他の材料:\n"
    for material, qty in other_materials.items():
        result += f"**{material}**  x  {int(qty)}\n"

    # クリスタルなどの特定の材料の表示
    result += f"\n必要なクリスタルの総量:\n"
    for material, qty in special_materials.items():
        result += f"{material}  x  {int(qty)}\n"

    # メッセージの送信
    await interaction.response.send_message(result)

@bot.tree.command(name='search_item', description='Search for a crafting item and display its materials')
@app_commands.describe(item_name='Name of the item to search for')
async def search_item(interaction: discord.Interaction, item_name: str):
    df = load_crafting_data()
    if df.empty:
        await interaction.response.send_message("エラー: Excelファイルを読み込む際に問題が発生しました。")
        return

    matching_items = df[df['完成品名'] == item_name]
    
    if not matching_items.empty:
        # アイテムが見つかった場合、その材料と必要数を返す
        response = f"アイテム '{item_name}' の材料リスト:\n"
        for _, row in matching_items.iterrows():
            response += f"  ・ {row['材料1']} x {int(row['必要数1'])}\n"
            if pd.notna(row['材料2']):
                response += f"  ・ {row['材料2']} x {int(row['必要数2'])}\n"
            if pd.notna(row['材料3']):
                response += f"  ・ {row['材料3']} x {int(row['必要数3'])}\n"
            if pd.notna(row['材料4']):
                response += f"  ・ {row['材料4']} x {int(row['必要数4'])}\n"
            if pd.notna(row['材料5']):
                response += f"  ・ {row['材料5']} x {int(row['必要数5'])}\n"
            if pd.notna(row['材料6']):
                response += f"  ・ {row['材料6']} x {int(row['必要数6'])}\n"
            if pd.notna(row['材料7']):
                response += f"  ・ {row['材料7']} x {int(row['必要数7'])}\n"
            if pd.notna(row['材料8']):
                response += f"  ・ {row['材料8']} x {int(row['必要数8'])}\n"
    else:
        # アイテムが見つからなかった場合、類似アイテムを提案
        all_items = df['完成品名'].tolist()
        similar_items = get_close_matches(item_name, all_items, n=10, cutoff=0.5)
        response = f"アイテム '{item_name}' が見つかりませんでした。類似するアイテム:\n" + "\n".join(similar_items)

    await interaction.response.send_message(response)

@bot.tree.command(name='mathelp', description='Display the list of available commands')
async def help_command(interaction: discord.Interaction):
    help_message = (
        "**RecipeResearcherBot Commands**\n"
        "/materials [材料名:数量,材料名:数量] - Calculate the total amount of materials needed.\n"
        "例: /materials 剛力の宝薬G2:9,魔匠の薬液:3\n"
        "/search_item [アイテム名] - Search for a crafting item and display its materials.\n\n"
    )
    await interaction.response.send_message(help_message)

def handle_terminal_commands():
    def terminal_commands():
        while True:
            command = input("コマンドを入力してください (reload_config, show_config, update_config <new_config>, exit): \n")
            if command.strip() == "reload_config":
                load_config()
                print(f"設定の読み込みなおしました。")
            elif command.strip() == "show_config":
                print(json.dumps(config, indent=4, ensure_ascii=False))
            elif command.startswith("update_config"):
                try:
                    new_config = json.loads(command[len("update_config"):].strip())
                    save_config(new_config)
                except Exception as e:
                    print(f"エラー: 設定の更新中に問題が発生しました - {e}")
            elif command.strip() == "exit":
                break
            else:
                print("無効なコマンドです。再度入力してください。")
    threading.Thread(target=terminal_commands, daemon=True).start()

if __name__ == "__main__":
    load_config()
    handle_terminal_commands()
    bot.run(config['DISCORD_BOT_TOKEN'])
