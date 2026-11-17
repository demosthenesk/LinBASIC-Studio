# LinBASIC Studio - the languages (LCID) of the Resource Editor
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""The language ids of Windows with the names Windows 7 gives them ("Greek (Greece)"), for the Language
lists of the Resource Editor (DEV-8/rs1), and the language of the system from the locale (el_GR -> 0x0408)."""
import locale
import os

LANGS = [
    (0x0436, "Afrikaans (South Africa)", "af_ZA"), (0x041C, "Albanian (Albania)", "sq_AL"),
    (0x0484, "Alsatian (France)", "gsw_FR"), (0x045E, "Amharic (Ethiopia)", "am_ET"),
    (0x1401, "Arabic (Algeria)", "ar_DZ"), (0x3C01, "Arabic (Bahrain)", "ar_BH"),
    (0x0C01, "Arabic (Egypt)", "ar_EG"), (0x0801, "Arabic (Iraq)", "ar_IQ"),
    (0x2C01, "Arabic (Jordan)", "ar_JO"), (0x3401, "Arabic (Kuwait)", "ar_KW"),
    (0x3001, "Arabic (Lebanon)", "ar_LB"), (0x1001, "Arabic (Libya)", "ar_LY"),
    (0x1801, "Arabic (Morocco)", "ar_MA"), (0x2001, "Arabic (Oman)", "ar_OM"),
    (0x4001, "Arabic (Qatar)", "ar_QA"), (0x0401, "Arabic (Saudi Arabia)", "ar_SA"),
    (0x2801, "Arabic (Syria)", "ar_SY"), (0x1C01, "Arabic (Tunisia)", "ar_TN"),
    (0x3801, "Arabic (U.A.E.)", "ar_AE"), (0x2401, "Arabic (Yemen)", "ar_YE"),
    (0x042B, "Armenian (Armenia)", "hy_AM"), (0x044D, "Assamese (India)", "as_IN"),
    (0x082C, "Azeri (Cyrillic, Azerbaijan)", "az_AZ@cyrillic"), (0x042C, "Azeri (Latin, Azerbaijan)", "az_AZ"),
    (0x046D, "Bashkir (Russia)", "ba_RU"), (0x042D, "Basque (Basque)", "eu_ES"),
    (0x0423, "Belarusian (Belarus)", "be_BY"), (0x0845, "Bengali (Bangladesh)", "bn_BD"),
    (0x0445, "Bengali (India)", "bn_IN"), (0x201A, "Bosnian (Cyrillic, Bosnia and Herzegovina)", "bs_BA@cyrillic"),
    (0x141A, "Bosnian (Latin, Bosnia and Herzegovina)", "bs_BA"), (0x047E, "Breton (France)", "br_FR"),
    (0x0402, "Bulgarian (Bulgaria)", "bg_BG"), (0x0403, "Catalan (Catalan)", "ca_ES"),
    (0x0C04, "Chinese (Traditional, Hong Kong S.A.R.)", "zh_HK"), (0x1404, "Chinese (Traditional, Macao S.A.R.)", "zh_MO"),
    (0x0804, "Chinese (Simplified, PRC)", "zh_CN"), (0x1004, "Chinese (Simplified, Singapore)", "zh_SG"),
    (0x0404, "Chinese (Traditional, Taiwan)", "zh_TW"), (0x0483, "Corsican (France)", "co_FR"),
    (0x041A, "Croatian (Croatia)", "hr_HR"), (0x101A, "Croatian (Latin, Bosnia and Herzegovina)", "hr_BA"),
    (0x0405, "Czech (Czech Republic)", "cs_CZ"), (0x0406, "Danish (Denmark)", "da_DK"),
    (0x048C, "Dari (Afghanistan)", "prs_AF"), (0x0465, "Divehi (Maldives)", "dv_MV"),
    (0x0813, "Dutch (Belgium)", "nl_BE"), (0x0413, "Dutch (Netherlands)", "nl_NL"),
    (0x0C09, "English (Australia)", "en_AU"), (0x2809, "English (Belize)", "en_BZ"),
    (0x1009, "English (Canada)", "en_CA"), (0x2409, "English (Caribbean)", "en_029"),
    (0x4009, "English (India)", "en_IN"), (0x1809, "English (Ireland)", "en_IE"),
    (0x2009, "English (Jamaica)", "en_JM"), (0x4409, "English (Malaysia)", "en_MY"),
    (0x1409, "English (New Zealand)", "en_NZ"), (0x3409, "English (Republic of the Philippines)", "en_PH"),
    (0x4809, "English (Singapore)", "en_SG"), (0x1C09, "English (South Africa)", "en_ZA"),
    (0x2C09, "English (Trinidad and Tobago)", "en_TT"), (0x0809, "English (United Kingdom)", "en_GB"),
    (0x0409, "English (United States)", "en_US"), (0x3009, "English (Zimbabwe)", "en_ZW"),
    (0x0425, "Estonian (Estonia)", "et_EE"), (0x0438, "Faroese (Faroe Islands)", "fo_FO"),
    (0x0464, "Filipino (Philippines)", "fil_PH"), (0x040B, "Finnish (Finland)", "fi_FI"),
    (0x080C, "French (Belgium)", "fr_BE"), (0x0C0C, "French (Canada)", "fr_CA"),
    (0x040C, "French (France)", "fr_FR"), (0x140C, "French (Luxembourg)", "fr_LU"),
    (0x180C, "French (Principality of Monaco)", "fr_MC"), (0x100C, "French (Switzerland)", "fr_CH"),
    (0x0462, "Frisian (Netherlands)", "fy_NL"), (0x0456, "Galician (Galician)", "gl_ES"),
    (0x0437, "Georgian (Georgia)", "ka_GE"), (0x0C07, "German (Austria)", "de_AT"),
    (0x0407, "German (Germany)", "de_DE"), (0x1407, "German (Liechtenstein)", "de_LI"),
    (0x1007, "German (Luxembourg)", "de_LU"), (0x0807, "German (Switzerland)", "de_CH"),
    (0x0408, "Greek (Greece)", "el_GR"), (0x046F, "Greenlandic (Greenland)", "kl_GL"),
    (0x0447, "Gujarati (India)", "gu_IN"), (0x0468, "Hausa (Latin, Nigeria)", "ha_NG"),
    (0x040D, "Hebrew (Israel)", "he_IL"), (0x0439, "Hindi (India)", "hi_IN"),
    (0x040E, "Hungarian (Hungary)", "hu_HU"), (0x040F, "Icelandic (Iceland)", "is_IS"),
    (0x0470, "Igbo (Nigeria)", "ig_NG"), (0x0421, "Indonesian (Indonesia)", "id_ID"),
    (0x085D, "Inuktitut (Latin, Canada)", "iu_CA"), (0x083C, "Irish (Ireland)", "ga_IE"),
    (0x0434, "isiXhosa (South Africa)", "xh_ZA"), (0x0435, "isiZulu (South Africa)", "zu_ZA"),
    (0x0410, "Italian (Italy)", "it_IT"), (0x0810, "Italian (Switzerland)", "it_CH"),
    (0x0411, "Japanese (Japan)", "ja_JP"), (0x044B, "Kannada (India)", "kn_IN"),
    (0x043F, "Kazakh (Kazakhstan)", "kk_KZ"), (0x0453, "Khmer (Cambodia)", "km_KH"),
    (0x0486, "K'iche (Guatemala)", "quc_GT"), (0x0487, "Kinyarwanda (Rwanda)", "rw_RW"),
    (0x0441, "Kiswahili (Kenya)", "sw_KE"), (0x0457, "Konkani (India)", "kok_IN"),
    (0x0412, "Korean (Korea)", "ko_KR"), (0x0440, "Kyrgyz (Kyrgyzstan)", "ky_KG"),
    (0x0454, "Lao (Lao P.D.R.)", "lo_LA"), (0x0426, "Latvian (Latvia)", "lv_LV"),
    (0x0427, "Lithuanian (Lithuania)", "lt_LT"), (0x082E, "Lower Sorbian (Germany)", "dsb_DE"),
    (0x046E, "Luxembourgish (Luxembourg)", "lb_LU"), (0x042F, "Macedonian (Former Yugoslav Republic of Macedonia)", "mk_MK"),
    (0x083E, "Malay (Brunei Darussalam)", "ms_BN"), (0x043E, "Malay (Malaysia)", "ms_MY"),
    (0x044C, "Malayalam (India)", "ml_IN"), (0x043A, "Maltese (Malta)", "mt_MT"),
    (0x0481, "Maori (New Zealand)", "mi_NZ"), (0x047A, "Mapudungun (Chile)", "arn_CL"),
    (0x044E, "Marathi (India)", "mr_IN"), (0x047C, "Mohawk (Mohawk)", "moh_CA"),
    (0x0450, "Mongolian (Cyrillic, Mongolia)", "mn_MN"), (0x0461, "Nepali (Nepal)", "ne_NP"),
    (0x0414, "Norwegian, Bokmål (Norway)", "nb_NO"), (0x0814, "Norwegian, Nynorsk (Norway)", "nn_NO"),
    (0x0482, "Occitan (France)", "oc_FR"), (0x0448, "Oriya (India)", "or_IN"),
    (0x0463, "Pashto (Afghanistan)", "ps_AF"), (0x0429, "Persian", "fa_IR"),
    (0x0415, "Polish (Poland)", "pl_PL"), (0x0416, "Portuguese (Brazil)", "pt_BR"),
    (0x0816, "Portuguese (Portugal)", "pt_PT"), (0x0446, "Punjabi (India)", "pa_IN"),
    (0x046B, "Quechua (Bolivia)", "quz_BO"), (0x0418, "Romanian (Romania)", "ro_RO"),
    (0x0417, "Romansh (Switzerland)", "rm_CH"), (0x0419, "Russian (Russia)", "ru_RU"),
    (0x243B, "Sami, Inari (Finland)", "smn_FI"), (0x103B, "Sami, Lule (Norway)", "smj_NO"),
    (0x043B, "Sami, Northern (Norway)", "se_NO"), (0x044F, "Sanskrit (India)", "sa_IN"),
    (0x1C1A, "Serbian (Cyrillic, Bosnia and Herzegovina)", "sr_BA@cyrillic"), (0x0C1A, "Serbian (Cyrillic, Serbia)", "sr_RS"),
    (0x181A, "Serbian (Latin, Bosnia and Herzegovina)", "sr_BA"), (0x081A, "Serbian (Latin, Serbia)", "sr_RS@latin"),
    (0x046C, "Sesotho sa Leboa (South Africa)", "nso_ZA"), (0x0432, "Setswana (South Africa)", "tn_ZA"),
    (0x045B, "Sinhala (Sri Lanka)", "si_LK"), (0x041B, "Slovak (Slovakia)", "sk_SK"),
    (0x0424, "Slovenian (Slovenia)", "sl_SI"), (0x2C0A, "Spanish (Argentina)", "es_AR"),
    (0x400A, "Spanish (Bolivia)", "es_BO"), (0x340A, "Spanish (Chile)", "es_CL"),
    (0x240A, "Spanish (Colombia)", "es_CO"), (0x140A, "Spanish (Costa Rica)", "es_CR"),
    (0x1C0A, "Spanish (Dominican Republic)", "es_DO"), (0x300A, "Spanish (Ecuador)", "es_EC"),
    (0x440A, "Spanish (El Salvador)", "es_SV"), (0x100A, "Spanish (Guatemala)", "es_GT"),
    (0x480A, "Spanish (Honduras)", "es_HN"), (0x080A, "Spanish (Mexico)", "es_MX"),
    (0x4C0A, "Spanish (Nicaragua)", "es_NI"), (0x180A, "Spanish (Panama)", "es_PA"),
    (0x3C0A, "Spanish (Paraguay)", "es_PY"), (0x280A, "Spanish (Peru)", "es_PE"),
    (0x500A, "Spanish (Puerto Rico)", "es_PR"), (0x0C0A, "Spanish (Spain, International Sort)", "es_ES"),
    (0x540A, "Spanish (United States)", "es_US"), (0x380A, "Spanish (Uruguay)", "es_UY"),
    (0x200A, "Spanish (Venezuela)", "es_VE"), (0x081D, "Swedish (Finland)", "sv_FI"),
    (0x041D, "Swedish (Sweden)", "sv_SE"), (0x045A, "Syriac (Syria)", "syr_SY"),
    (0x0428, "Tajik (Cyrillic, Tajikistan)", "tg_TJ"), (0x085F, "Tamazight (Latin, Algeria)", "tzm_DZ"),
    (0x0449, "Tamil (India)", "ta_IN"), (0x0444, "Tatar (Russia)", "tt_RU"),
    (0x044A, "Telugu (India)", "te_IN"), (0x041E, "Thai (Thailand)", "th_TH"),
    (0x0451, "Tibetan (PRC)", "bo_CN"), (0x041F, "Turkish (Turkey)", "tr_TR"),
    (0x0442, "Turkmen (Turkmenistan)", "tk_TM"), (0x0422, "Ukrainian (Ukraine)", "uk_UA"),
    (0x042E, "Upper Sorbian (Germany)", "hsb_DE"), (0x0420, "Urdu (Islamic Republic of Pakistan)", "ur_PK"),
    (0x0480, "Uyghur (PRC)", "ug_CN"), (0x0843, "Uzbek (Cyrillic, Uzbekistan)", "uz_UZ@cyrillic"),
    (0x0443, "Uzbek (Latin, Uzbekistan)", "uz_UZ"), (0x042A, "Vietnamese (Vietnam)", "vi_VN"),
    (0x0452, "Welsh (United Kingdom)", "cy_GB"), (0x0488, "Wolof (Senegal)", "wo_SN"),
    (0x0485, "Yakut (Russia)", "sah_RU"), (0x0478, "Yi (PRC)", "ii_CN"), (0x046A, "Yoruba (Nigeria)", "yo_NG"),
]
_BY_ID = {i: n for i, n, _ in LANGS}
_BY_LOC = {c: i for i, _, c in LANGS}


def name(lcid):
    """the name Windows gives a language, or its number (as VB6 lists the ones without a name)"""
    return _BY_ID.get(lcid, str(lcid))


def names():
    """[(lcid, name)] sorted as the Language lists of VB6 show them: by name"""
    return sorted(((i, n) for i, n, _ in LANGS), key=lambda x: x[1].upper())


# the languages VB6 lists by number only, before the names (DEV-8/rs1 stage 7, picture 26): the ids its table
# knows and Windows 7 gives no name; ascending, as VB6 shows them
NUMERIC = [1112, 1113, 2080, 2087, 2144, 2145]


def combo_items():
    """[(lcid, text)] of a Language list of the Resource Editor: the numbers first, then the names"""
    return [(i, str(i)) for i in NUMERIC] + names()


def from_name(text):
    for i, n, _ in LANGS:
        if n == text:
            return i
    try:
        return int(text)
    except (TypeError, ValueError):
        return None


def system_lcid():
    """the language of the system (LANG, LC_ALL ...): el_GR.UTF-8 -> 0x0408; English (United States) else"""
    for var in ("LC_ALL", "LC_MESSAGES", "LANG"):
        v = os.environ.get(var)
        if v:
            break
    else:
        v = (locale.getlocale()[0] or "")
    code = v.split(".")[0]
    if code in _BY_LOC:
        return _BY_LOC[code]
    lang = code.split("_")[0]
    for i, _, c in LANGS:
        if c.split("_")[0] == lang:
            return i
    return 0x0409
