from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
import logging
import os
from suivi import Suivi
from excel_fonctions import Excel
import dataframe_image as dfi
from vendeur import vendeur_number_phone

TOKEN = "6419312810:AAG4Zt2jP9SjkB80I5hSEatlulmOSE4zB7E"
logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
dp = Dispatcher(bot)
code_vendeurs = [i.split(" ")[0] for i in vendeur_number_phone.keys()]
name_vendeurs = [i.split(" ", 1)[1] for i in vendeur_number_phone.keys()]
code_name_vendeurs = list(vendeur_number_phone.keys())
button_vendeur = KeyboardButton("Vendeur")
button_admin = KeyboardButton("Admin")

btn_obj = KeyboardButton("Obj")
btn_Quantitatif = KeyboardButton("Quantitatif")
btn_Qualitatif = KeyboardButton("Qualitatif")
btn_focus = KeyboardButton("Focus")
btn_send_excel = KeyboardButton("Send excel")

keyboard_admin_vendeur = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).row(button_vendeur,
                                                                                               button_admin)
keyboard_vendeur = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).row(btn_Quantitatif,
                                                                                         btn_Qualitatif, btn_obj,
                                                                                         btn_focus)
keyboard_admin = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).row(btn_send_excel, btn_Quantitatif,
                                                                                       btn_Quantitatif)

code_vendeur = ""
name_vendeur = ""


def search_index(list: list, name: str) -> int:
    for i in list:
        if i.startswith(name):
            return list.index(i)


@dp.message_handler(commands=["start", "help"])
async def welcome(message: types.Message):
    await message.answer("Hi are you Admin or Vendeur", reply_markup=keyboard_admin_vendeur)


@dp.message_handler()
async def echo(message: types.Message):
    if message.text == "Vendeur":
        await message.answer(f'Saisie votre code vendeur')
    if message.text in code_vendeurs:
        global code_vendeur
        code_vendeur = message.text
        global index
        index = search_index(code_name_vendeurs, message.text)
        global name_vendeur
        name_vendeur = name_vendeurs[index]

        await message.answer(f'Bonjour {name_vendeurs[index]} \n tu veux quoi?', reply_markup=keyboard_vendeur)
    if message.text == "Quantitatif":
        await send_quantitatif(message.chat.id, f'{code_vendeur} {name_vendeur}')
    if message.text == "Qualitatif":
        await send_qualitatif(message.chat.id, f'{code_vendeur} {name_vendeur}')

    if message.text == "Admin":
        await message.answer(f'Saisie votre mot de pass')
    if message.text == "iamsorry":
        await message.answer(f'Bonjour Chakib', reply_markup=keyboard_admin)
    if message.text == "Send excel":
        await message.answer(f'Send a excel file', )


async def send_qualitatif(chat_id, vendeur_name: str, month: int = 24, days_work: int = 1):
    suivi = Suivi(vendeur_name=vendeur_name, total_to_work=month, days_work=days_work)
    suivi.send_qualitatif()
    await bot.send_photo(chat_id=chat_id, photo=open(f"excel/{code_vendeur} {name_vendeur}.jpeg", 'rb'),
                         caption="jaune : TTC",
                         reply_markup=keyboard_vendeur)
    os.remove(f"excel/{code_vendeur} {name_vendeur}.jpeg")


async def send_quantitatif(chat_id, vendeur_name: str, month: int = 24, days_work: int = 1):
    suivi = Suivi(vendeur_name=vendeur_name, total_to_work=month, days_work=days_work)
    suivi.send_quantitatif()
    await bot.send_photo(chat_id=chat_id, photo=open(f"excel/{code_vendeur} {name_vendeur}.jpeg", 'rb'),
                         caption="jaune : TTC",
                         reply_markup=keyboard_vendeur)
    os.remove(f"excel/{code_vendeur} {name_vendeur}.jpeg")


@dp.message_handler(content_types=types.ContentTypes.DOCUMENT)
async def handle_document(message: types.Message):
    # Get the file ID from the received message
    file_id = message.document.file_id

    # Get the file object using the file ID
    file = await bot.get_file(file_id)

    # Download the file
    file_path = file.file_path
    downloaded_file = await bot.download_file(file_path)

    # Specify the path where you want to save the downloaded file
    save_path = f'excel/suivi.xlsx'

    # Save the file to the specified path
    with open(save_path, 'wb') as new_file:
        new_file.write(downloaded_file.read())

        # Send a response to the user indicating successful file saving
        await message.reply(f"File '{message.document.file_name}' saved successfully!")
        # Make a mod file
        a = Excel("excel/suivi.xlsx")

        a.fix_sheets()

        w_day, t_day = a.get_day_work()
        d_rest = t_day - w_day


print("Program start")
executor.start_polling(dp)
