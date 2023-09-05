from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
import logging
import json
from suivi import Suivi
from core import MyFonction
from excel_fonctions import Excel
import os
from vendeur import vendeur_number_phone


def main():
    TOKEN = "6419312810:AAFaGXyhKZMmA48T0lk3HnZVu-yaBmnAvrw"
    logging.basicConfig(level=logging.INFO)
    bot = Bot(token=TOKEN)
    dp = Dispatcher(bot)

    code_name_vendeurs: list = list(vendeur_number_phone.keys())
    code_vendeurs = [i.split(" ")[0] for i in vendeur_number_phone.keys()]

    button_vendeur = KeyboardButton("Vendeur")
    button_admin = KeyboardButton("Admin")

    btn_obj = KeyboardButton("Obj")
    btn_Quantitatif = KeyboardButton("Quantitatif")
    btn_Qualitatif = KeyboardButton("Qualitatif")
    btn_focus = KeyboardButton("Focus")
    btn_send_excel = KeyboardButton("Send excel")

    keyboard_admin_vendeur = ReplyKeyboardMarkup(
        resize_keyboard=True, one_time_keyboard=True
    ).row(button_vendeur, button_admin)
    keyboard_vendeur = ReplyKeyboardMarkup(
        resize_keyboard=True, one_time_keyboard=True
    ).row(btn_Quantitatif, btn_Qualitatif, btn_obj, btn_focus)

    def search_index(my_list: list, name: str) -> int:
        for i in my_list:
            if i.startswith(name):
                return my_list.index(i)

    @dp.message_handler(commands=["start", "help"])
    async def welcome(message: types.Message):
        await message.answer(
            "Hi are you Admin or Vendeur", reply_markup=keyboard_admin_vendeur
        )

    def write_file(value: str):
        # Open a file in write mode
        try:
            data = {"days": {"rr": value}}

            # Specify the file name
            file_name = "days.json"

            # Open the file in write mode and use json.dump() to write the data as JSON
            with open(file_name, "w") as json_file:
                json.dump(data, json_file, indent=4)
                print("Data written successfully.")
        except Exception as e:
            print("An error occurred:", e)

    async def send_qualitatif(chat_id, vendeur_name: str):
        suivi = Suivi(vendeur_name)
        suivi.save_qualitatif_image()
        await bot.send_photo(
            chat_id=chat_id,
            photo=open(f"excel/{vendeur_name}.jpeg", "rb"),
            caption="jaune : TTC",
            reply_markup=keyboard_vendeur,
        )
        os.remove(f"excel/{vendeur_name}.jpeg")

    async def send_quantitatif(
            chat_id,
            vendeur_name: str,
    ):
        print(vendeur_name)
        suivi = Suivi(vendeur_name)
        suivi.save_quantitatif_image()
        await bot.send_photo(
            chat_id=chat_id,
            photo=open(f"excel/{vendeur_name}.jpeg", "rb"),
            caption="jaune : TTC",
            reply_markup=keyboard_vendeur,
        )
        os.remove(f"excel/{vendeur_name}.jpeg")

    @dp.message_handler()
    async def echo(message: types.Message):
        global code_name_vendeur
        if message.text == "Vendeur":
            await message.answer(f"Saisie votre code vendeur")
        if message.text in code_vendeurs:
            code_name_vendeur = MyFonction.get_code_name(
                code_name_vendeurs, message.text
            )
            await message.answer(
                f"Bonjour {code_name_vendeur.split(' ', 1)[1]} \n tu veux quoi?",
                reply_markup=keyboard_vendeur,
            )
        if message.text == "Quantitatif":
            await send_quantitatif(message.chat.id, f"{code_name_vendeur}")
        if message.text == "Qualitatif":
            await send_qualitatif(message.chat.id, f"{code_name_vendeur}")

        if message.text == "Admin":
            await message.answer(f"Saisie votre mot de pass")
        if message.text == "iamsorry":
            await message.answer(f"Bonjour Chakib. write real days rest")
        if message.text[0].isdigit():
            write_file(int(message.text))
            await message.answer("Send Excel now!")

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
        save_path = f"excel/suivi.xlsx"

        # Save the file to the specified path
        with open(save_path, "wb") as new_file:
            new_file.write(downloaded_file.read())

            # Send a response to the user indicating successful file saving
            await message.reply(
                f"File '{message.document.file_name}' saved successfully!"
            )
            # Make a mod file
            a = Excel("excel/suivi.xlsx")

            a.fix_sheet()

            w_day, t_day = a.get_day_work()

            with open("days.json", "r") as jsonFile:
                data = json.load(jsonFile)

            data["from_file"] = {"t": t_day, "d": w_day}

            with open("days.json", "w") as jsonFile:
                json.dump(data, jsonFile)

    print("Program start")
    executor.start_polling(dp)


if __name__ == "__main__":
    main()
