import discord
from openpyxl import Workbook
import xlrd

pdfs_dict = {}
pdfs_list=[]
def read_data_from_xlsx_file(user_message):
    workbook = xlrd.open_workbook("Tickers.xlsx", "rb")
    sheets = workbook.sheet_names()
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)
        for rownum in range(1, sh.nrows):
            pdfs_list=[]
            row_valaues = sh.row_values(rownum)
            pdfs_list.append(row_valaues[1])
            pdfs_list.append(row_valaues[3])
            pdfs_dict[f"{row_valaues[0]}"]=pdfs_list

    return str(user_message)

async def send_message(message, user_message):
    try:
        response = read_data_from_xlsx_file(user_message)

        # await message.author.send(response) if is_private else await message.channel.send(response)
        # await message.channel.send(file=discord.File(response))

        if str(response).split()[0].lower()=='bot':
            print(message.channel)
            response=pdfs_dict.get(str(response).split(' ')[-1])
            if response==None: await message.channel.send('Please enter a correct ticker !')

            else:
                if str(response[1]).lower()==str(message.channel).replace('-',' '):
                    # await message.channel.send(str(response[0]))
                    await message.channel.send(file=discord.File(f'pdf_files/{str(response[0])}'))
                else:
                    await message.channel.send(f'Sorry this ticker belongs to {str(response[1])} sector I can not send it !!')

    except Exception as e:
        print(e)

def run_discord_bot():
    TOKEN = ''  #TODO : ENTER YOUR TOKEN HERE
    intents = discord.Intents.default()
    intents.message_content = True
    client = discord.Client(intents=intents)

    @client.event
    async def on_ready():
        print(f'{client.user} is now running!')

    @client.event
    async def on_message(message):
        if message.author==client.user:
            return

        username=str(message.author)
        user_message=str(message.content)
        channel=str(message.channel)

        print(f'{username} said: "{user_message}" "({channel})')

        await send_message(message,user_message)

    client.run(TOKEN)

