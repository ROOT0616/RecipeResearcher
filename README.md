
# RecipeResearcherBot

RecipeResearcherBot is a Discord bot designed to help players of Final Fantasy XIV (FF14) calculate crafting materials needed for various items. The bot uses an Excel file to retrieve crafting recipes and outputs the total number of materials required based on user input.

## Features

- **/materials [item_name:quantity]**: Calculates the total amount of materials needed for the specified items and quantities.
- **/search_item [item_name]**: Searches for a crafting item and displays its materials (up to 8 materials supported).
- **/mathelp**: Displays the list of available commands.
- Supports up to 8 materials per item.
- Logs bot activities including errors and command usage to a log file.
- Handles Excel file reloading and configuration updates via terminal commands.

## Usage

### Discord Commands
- **/materials [item_name:quantity]**: Calculates the total amount of materials needed for specified items. Example: `/materials 剛力の宝薬G2:9,魔匠の薬液:3`.
- **/search_item [item_name]**: Searches for a crafting item and displays its materials. Example: `/search_item クラロウォルナット・スピア`.
- **/mathelp**: Displays help information about the available commands.

### Terminal Commands
- **reload_config**: Reloads the bot's configuration from the `config.json` file.
- **show_config**: Displays the current configuration in JSON format.
- **update_config <new_config>**: Updates the `config.json` with new configuration values.
- **exit**: Shuts down the bot.

## Installation

1. Install the required Python packages using `pip`:
   ```bash
   pip install discord.py pandas openpyxl
   ```
2. Add your bot token to the `config.json` file.
3. Run the bot using:
   ```bash
   python RecipeResearcherBot.py
   ```

## Configuration

The bot's configuration file, `config.json`, contains the following settings:
- **EXCEL_FILE_PATH**: Path to the Excel file that contains crafting recipes.
- **SPECIAL_ITEMS**: List of special items (e.g., crystals) to handle separately.
- **DISCORD_BOT_TOKEN**: Your Discord bot token.

## License
This project is licensed under the MIT License.
