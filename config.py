import os

BASE_FOLDER = r"C:\Users\VISSERLU\Anheuser-Busch InBev\SFDC COE - EUR - General\Display Load"


COUNTRY_RULES = {
    "BE data": {
        "keywords": ["beoff.displays@ab-inbev.com", "Mickey"],
        "prefix": "ABI - Display Data",
        "folder": os.path.join(BASE_FOLDER, "BE Display Load"),
        "country_code": "BE",
    },
    "FR": {
        "keywords": ["cpm france", "stacy.balounaik@cpm-int.com", "Billancourt"],
        "prefix": "Tracking Display",
        "folder": os.path.join(BASE_FOLDER, "FR Display Load"),
        "country_code": "FR",
    },
    "NL": {
        "keywords": ["the netherlands", "hamiltonbright.com"],
        "prefix": "NL Display",
        "folder": os.path.join(BASE_FOLDER, "NL Display Load"),
        "country_code": "NL",
    },
    "IT": {
        "keywords": ["milano",
                     "Italy",
                     "Italia",],
        "prefix": "IT Display",
        "folder": os.path.join(BASE_FOLDER, "IT Display Load"),
        "country_code": "IT",
    },
    "BE Week": {
        "keywords": [
            "beoff.displays@ab-inbev.com",
            "impactsalesmarketing.be",
            "g.vangelder@impactfieldmarketinggroup.com",
            "BE",
            "Belgium",
            "Belgique",
            "Belg",
        ],
        "prefix": "BE Weekly Report",
        "folder": os.path.join(BASE_FOLDER, "BE Display Load"),
        "country_code": "BE",
    },
}


SNIPPET_LEN = 1000
MAX_EMAILS = 50
EMAIL_LOOKBACK_DAYS = 1
DEBUG = True


DEFAULT_DISPLAY_ID = "a2e24000000TuGw"


KEEP_COLS = [
    "ABI_SFA_POC__C",
    "ABI_SFA_PRODUCT_SET__C",
    "ABI_SFA_QUANTITY_IN_SURVEY__C",
    "ABI_SFA_START_DATE__C",
    "ABI_SFA_END_DATE__C",
    "ABI_SFA_NUMBER_OF_WEEK__C",
    "ABI_SFA_DISPLAY_SEQUENCE__C",
    "ABI_SFA_DISPLAY_TYPE__C",
    "ABI_SFA_DISPLAY__C",
    "ABI_SFA_FOR_BATCH_PROCESSING__C",
    "ABI_SFA_LOCATION__C",
    "ABI_SFA_PERSON_REGISTERED__C",
    "ABI_SFA_STATUS__C",
    "ABI_SFA_MECHANISM__C",
    "ABI_SFA_EXTERNAL_ID__C",
    "ABI_SFA_POCM__C",
    "ABI_SFA_CONDITION_NAME__C",
    "ABI_SFA_M1__C",
]
