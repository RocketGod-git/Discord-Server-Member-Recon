import discord
from discord import Intents
import pandas as pd
from openpyxl import Workbook
from urllib.parse import urlparse, urlunparse
import importlib
import subprocess
import traceback
import sys

# Step 1: Create a Discord application and bot account, and get the bot token
bot_token = 'put your discord bots token inside here'
channel_name = 'put your channel name here'  # Change to your desired channel name to send progress and error handling to. Make sure the bot has permission to write there!

# Step 2: Check if necessary packages (discord.py, pandas, openpyxl) are installed
required_modules = ['discord', 'pandas', 'openpyxl']
for module in required_modules:
    try:
        importlib.import_module(module)
        print(f"Module {module} is installed.")
    except ImportError:
        print(f"{module} is not installed. Installing...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', module])

# Step 3: Connect to the Discord API and fetch member information
intents = Intents.default()
intents.members = True

client = discord.Client(intents=intents)

async def send_embedded_message(title, description, color=0x00ff00):
    channel = discord.utils.get(client.get_all_channels(), name=channel_name)
    if channel:
        embed = discord.Embed(title=title, description=description, color=color)
        await channel.send(embed=embed)

@client.event
async def on_ready():
    try:
        print(f"Successfully logged into Discord as {client.user.name}")
                
        guild = client.guilds[0]  # Select the first guild the bot is in
        members = guild.members

        # Extract member information
        data = {'Name': [], 'Discriminator': [], 'Nickname': [], 'Avatar_URL': [], 'ID': [], 'Roles': [], 'Top_Role': [], 'Joined_at': [], 'Created_at': [], 'Bot': [], 'Status': [], 'Activity': [], 'Desktop_Status': [], 'Mobile_Status': [], 'Web_Status': [], 'Raw_Status': []}
        for member in members:
            data['Name'].append(member.name)
            data['Discriminator'].append(member.discriminator)
            data['Nickname'].append(member.nick)
            if member.avatar is not None:
                avatar_url = member.avatar.url
                if not avatar_url.startswith('http'):
                    if avatar_url.startswith('a_'):
                        avatar_url = f'https://cdn.discordapp.com/avatars/{member.id}/{avatar_url}.gif'
                    else:
                        avatar_url = f'https://cdn.discordapp.com/avatars/{member.id}/{avatar_url}.png'
            else:
                avatar_url = None
            data['Avatar_URL'].append(avatar_url)
            data['ID'].append(member.id)
            roles = [role.name for role in member.roles]
            data['Roles'].append(roles)
            top_role = member.top_role.name if member.top_role.name != "@everyone" else " "
            data['Top_Role'].append(top_role)
            data['Joined_at'].append(str(member.joined_at))
            data['Created_at'].append(str(member.created_at))
            data['Bot'].append(member.bot)
            data['Status'].append(str(member.status))
            data['Activity'].append(str(member.activity))
            data['Desktop_Status'].append(str(member.desktop_status))
            data['Mobile_Status'].append(str(member.mobile_status))
            data['Web_Status'].append(str(member.web_status))
            data['Raw_Status'].append(str(member.raw_status))

        print("Information fetched successfully.")
        await send_embedded_message("Progress", "Information fetched successfully.")

        # Step 4: Organize the fetched information into a spreadsheet using the Pandas library
        df = pd.DataFrame(data)
        print("Data organized into a DataFrame.")
        await send_embedded_message("Progress", "Data organized into a DataFrame.")

        # Step 5: Export the spreadsheet to an Excel file using the openpyxl library
        excel_file_name = 'discord_members.xlsx'
        with pd.ExcelWriter(excel_file_name) as writer:
            df.to_excel(writer, index=False, sheet_name='members')
        print("Data exported to 'discord_members.xlsx'.")
        await send_embedded_message("Progress", "Data exported to 'discord_members.xlsx'.")

        # Step 6: Export an HTML file with embedded avatar images
        html_file_name = 'discord_members.html'
        def embed_image(url):
            if url:
                if url.startswith('https://cdn.discordapp.com/avatars/'):
                    url = url.split("?")[0]  # Remove query parameter from avatar URL
                return f'<img src="{url}" style="width: 50px; height:50px;">'
            else:
                return ""

        df["Avatar_URL"] = df["Avatar_URL"].apply(embed_image)
        html = df.to_html(escape=False, index=False)
        with open(html_file_name, 'w', encoding='utf-8') as f:
            f.write(html)
        print("HTML file 'discord_members.html' created.")
        await send_embedded_message("Progress", "HTML file 'discord_members.html' created.")
        
        # Step 7: Send files to the Discord channel
        channel = discord.utils.get(client.get_all_channels(), name=channel_name)
        if channel:
            await channel.send(file=discord.File(excel_file_name))
            await channel.send(file=discord.File(html_file_name))
            print("Files sent to the Discord channel.")
                    
        # Step 8: Send a custom message to the Discord channel
        await channel.send("All done, have a great day! -RocketGod")
        print("Custom message sent to the Discord channel.")
                
        await client.close()

    except Exception as e:
        print(f"An error occurred: {e}")
        traceback.print_exc()
        await send_embedded_message("Error", str(e), color=0xff0000)

try:
    client.run(bot_token)
except discord.LoginFailure:
    print(f"Failed to log into Discord. Check if the provided bot token ({bot_token}) is correct.")
except Exception as e:
    print(f"An error occurred: {e}")
    traceback.print_exc()