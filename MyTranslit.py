from transliterate.base import TranslitLanguagePack


class MyTranslit(TranslitLanguagePack):
    language_code = "ru_nal"
    language_name = "ru_naladka"
    mapping = (
        u"АБВГДЕЖЗИКЛМНОПРСТУФХЦЩЫЮабвгдежзиклмнопрстуфхцщыю",
        u"ABVGDEJZIKLMNOPRSTUFXCHYQabvgdejziklmnoprstufxchyq"
    )
    pre_processor_mapping = {
        # uppercase
        u"Ч": u"CH",
        u"Ш": u"W",
        u"Ъ": u"jj",
        u"Ь": u"ii",
        u"Э": u"Je",
        u"Я": u"Ja",

        # lowercase
        u"ч": u"ch",
        u"ш": u"w",
        u"ъ": u"jj",
        u"ь": u"ii",
        u"э": u"je",
        u"я": u"ja"
    }
