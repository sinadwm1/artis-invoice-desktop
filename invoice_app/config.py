CELL_MAPPING = {
    "issue_date": "M3",
    "validation_date": "M4",
    "invoice_no": "M5",

    "buyer_company_name": "L9",
    "buyer_national_id": "L10",
    "buyer_address": "K11",
    "buyer_website": "K12",
    "buyer_phone": "K13",
    "buyer_fax": "K14",

    "shipping_cost": "K35",
    "grand_total": "K36",
    "term_of_payment": "F38",

    "shipping_mode_of_transport": "F42",
    "shipping_terms_of_delivery": "F43",
    "shipping_load_date": "F44",
    "shipping_partial": "F45",
    "shipping_beneficiary_country": "F46",
    "shipping_origin_country": "F47",
    "shipping_weight_per_unit": "E48",
    # تغییر به E49 طبق درخواست شما:
    "shipping_dimensions": "F49",

    "bank_name": "E44",
    "bank_swift": "E45",
    "bank_iban": "E46",
    "bank_account": "E47",

    "custom_hs_code": "F16"
}

ITEM_MAPPING = {
    "description": "D",
    "quantity": "K",
    "unit_price": "L",
    "total": "N"
}

ITEM_ROWS = {index: 18 + index for index in range(1, 17)}

MAX_ITEMS = 16

TEMPLATES = {
    "Blue Template": "blue_template.xlsx",
    "Red Template": "red_template.xlsx"
}
