import discord
from discord import Intents
import pandas as pd
from openpyxl import Workbook
from IPython.display import HTML
import base64
from urllib.parse import urlparse, urlunparse
import importlib
import subprocess

# Step 1: Create a Discord application and bot account, and get the bot token
bot_token = 'INSERT YOUR BOT TOKEN HERE'

# Step 2: Check if necessary packages (discord.py, pandas, openpyxl) are installed
required_modules = ['discord', 'pandas', 'openpyxl']
for module in required_modules:
    try:
        importlib.import_module(module)
    except ImportError:
        print(f"{module} is not installed. Installing...")
        subprocess.call(['pip', 'install', module])

# Step 3: Connect to the Discord API and fetch member information
intents = Intents.default()
intents.members = True

client = discord.Client(intents=intents)

@client.event
async def on_ready():
    guild = client.guilds[0] # Select the first guild the bot is in
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

    # Step 4: Organize the fetched information into a spreadsheet using the Pandas library
    df = pd.DataFrame(data)

    # Step 5: Export the spreadsheet to an Excel file using the openpyxl library
    with pd.ExcelWriter('discord_members.xlsx') as writer:
        df.to_excel(writer, index=False, sheet_name='members')

    # Step 6: Export an HTML file with embedded avatar images
    def embed_image(url):
        if url:
            if url.startswith('https://cdn.discordapp.com/avatars/'):
                url = url.split("?")[0]  # Remove query parameter from avatar URL
            return f'<img src="{url}" style="width: 50px; height:50px;">'
        else:
            return ""

    df["Avatar_URL"] = df["Avatar_URL"].apply(embed_image)
    html = df.to_html(escape=False, index=False)
    with open('discord_members.html', 'w', encoding='utf-8') as f:
        f.write(html)

    
    await client.close()

client.run(bot_token)
