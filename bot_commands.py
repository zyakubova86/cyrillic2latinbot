from aiogram import types


async def set_default_commands(dp):
    await dp.bot.set_my_commands([
        types.BotCommand('start', "Start"),
        types.BotCommand('cancel', "Cancel"),
        types.BotCommand('cyrillic2latin', "Cyrillic <> latin"),
        types.BotCommand('ivmsfile', "IVMS"),
        types.BotCommand('qr', "QR code"),
        types.BotCommand('getmyid', "ID"),
    ])
