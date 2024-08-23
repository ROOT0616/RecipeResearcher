import discord
from discord import app_commands
from discord.ext import commands
import pandas as pd
import math
import json
from difflib import get_close_matches

# config.json ファイルから設定を読み込む
with open('config.json', 'r', encoding='utf-8') as config_file:
    config = json.load(config_file)

# Discordトークンを設定
TOKEN = config['DISCORD_BOT_TOKEN']

# Botの設定
intents = discord.Intents.default()
intents.message_content = True
bot = commands.Bot(command_prefix='/', intents=intents)

# Excelファイルのパス
EXCEL_FILE_PATH = config['EXCEL_FILE_PATH']

# 特定のアイテムリスト
SPECIAL_ITEMS = config['SPECIAL_ITEMS']

def load_crafting_data():
    try:
        # Excelファイルを読み込み
        df = pd.read_excel(EXCEL_FILE_PATH)
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
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
        "Please make sure to format your input as 'Material:Quantity,Material:Quantity'."
    )
    await interaction.response.send_message(help_message)

# ボットの起動
bot.run(TOKEN)
