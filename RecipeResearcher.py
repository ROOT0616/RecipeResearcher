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
        # Excelファイルを読み込み
        df = pd.read_excel(EXCEL_FILE_PATH)
        return df
    except Exception as e:
        logging.error(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
        return pd.DataFrame()

def calculate_materials(materials, df):
    total_materials = {}

    for item, quantity in materials.items():
        item_data = df[df['完成品名'] == item]

        if item_data.empty:
            continue

        for _, row in item_data.iterrows():
            completion_quantity = row['完成個数']  # 完成品の1回あたりの数

            for col in df.columns:
                if '材料' in col:
                    material = row[col]
                    
                    if isinstance(material, str):  # 文字列の場合のみ処理
                        material = material.strip()
                    elif pd.notna(material):
                        material = str(material)  # 数値を文字列に変換
                    else:
                        continue  # NaN や空白の場合はスキップ
                    
                    required_quantity = row[col.replace('材料', '必要数')]
                    total_quantity = quantity * required_quantity
                    recipe_count = math.ceil(total_quantity / completion_quantity)

                    if material not in total_materials:
                        total_materials[material] = 0
                    total_materials[material] += recipe_count * required_quantity

    final_materials = {}
    for material, qty in total_materials.items():
        if material in df['完成品名'].values:
            nested_materials = calculate_materials({material: qty}, df)
            for nested_material, nested_qty in nested_materials.items():
                if nested_material not in final_materials:
                    final_materials[nested_material] = 0
                final_materials[nested_material] += nested_qty
        else:
            if material not in final_materials:
                final_materials[material] = 0
            final_materials[material] += qty

    return final_materials

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
        for arg in materials.split(','):
            item, qty = arg.split(':')
            materials_dict[item] = int(qty)
    except Exception as e:
        await interaction.response.send_message("エラー: コマンドの形式が正しくありません。例: /materials 剛力の宝薬G2:9,魔匠の薬液:3")
        logging.error(f"材料コマンドの形式エラー: {e}")
        return

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

    total_materials = calculate_materials(materials_dict, df)
    
    if not total_materials:
        await interaction.response.send_message("エラー: 指定された材料に基づくデータが見つかりませんでした。")
        return

    # 特定のアイテムをリストの最後に持ってくるための処理
    other_materials = {material: qty for material, qty in total_materials.items() if material not in SPECIAL_ITEMS}
    special_materials = {material: qty for material, qty in total_materials.items() if material in SPECIAL_ITEMS}

    # 文字列を行ごとに分割
    lines = materials.split(',')

    # 辞書を初期化
    material_dict = {}
    
    # 各行を分解して辞書に追加
    for line in lines:
        # 空白を取り除き、':'で分割
        key, value = line.strip().split(':')
        # 辞書に追加
        material_dict[key] = int(value)
    
    # 結果のフォーマット
    materialss = ""
    for material, quantity in material_dict.items():
        materialss += f"**{material}**  を  {quantity}個\n"

    result = f"{materialss}の必要な材料の総量:\n\n"
    
    for material, qty in other_materials.items():
        result += f"**{material}**  x  {int(qty)}\n"

    result += f"\n必要なクリスタルの総量:\n"
    for material, qty in special_materials.items():
        result += f"{material}  x  {int(qty)}\n"
    
    await interaction.response.send_message(result)

@bot.tree.command(name='mathelp', description='Display the list of available commands')
async def help_command(interaction: discord.Interaction):
    help_message = (
        "**RecipeResearcherBot Commands**\n"
        "/materials [材料名:数量,材料名:数量] - Calculate the total amount of materials needed.\n"
        "例: /materials 剛力の宝薬G2:9,魔匠の薬液:3\n\n"
        "このコマンドを使用すると、指定した材料の総量を計算し、結果を返します。\n"
        "\n"
        "**Terminal Commands**\n"
        "reload_config - Reload the configuration file.\n"
        "show_config - Display the current configuration.\n"
        "update_config <new_config> - Update the configuration with the provided JSON.\n"
        "exit - Exit the terminal command interface.\n"
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
