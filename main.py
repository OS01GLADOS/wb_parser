import requests

import openpyxl
from openpyxl import Workbook


initial_address = "https://static-basket-01.wbbasket.ru/vol0/data/main-menu-by-ru-v3.json"

item_list_query = "https://search.wb.ru/exactmatch/ru/common/v14/search?ab_testing=false&appType=1&curr=byn&dest=-1257786&hide_dtype=10;13;14&lang=ru&page=1&query={category_value}&resultset=catalog&sort=popular&spp=30&suppressSpellcheck=false"

ITEM_NESTING_LV = 99

OUTPUT_FILE_NAME = 'wb.xlsx'

item_data_all = {}


def get_data_by_request(address):
    response = requests.get(address)
    data = response.json()
    return data


def get_items(id):
    address = item_list_query.format(category_value=id)
    items = get_data_by_request(address)
    return items


def get_childs(category_id):
    items = get_items(category_id)
    return parse_items(items)


def parse_childs(children, nesting_lv=0):
    nesting_lv += 1
    categories = {}
    for child in children:
        id = child["id"]
        name = child["name"]
        url = child["url"]
        query=child.get("searchQuery", None)
        children_inner = child.get("childs")
        item_to_append = {"name": name, "id": id, "nesting_lv": nesting_lv, "url": url, 'query': query}
        if children_inner:
            childs = parse_childs(children_inner, nesting_lv)
            item_to_append["children"] = childs
        categories[id] = item_to_append
    return categories


def parse_items(item):
    items = {}
    products = item.get("products")
    for product in products:
        id = product["id"]
        brand = product["brand"]
        name = product["name"]
        colors = product["colors"]
        nesting_lv = ITEM_NESTING_LV
        element_to_append = {"name": name, "brand": brand, "colors": colors, "nesting_lv": nesting_lv}
        items[id] = element_to_append
    return items


def load_childs(categories):
    all_count = len(categories)
    remaining_count = len(categories)
    for category_id, query in categories.items():
        remaining_count -=1
        print(f"loading {category_id}(query: {query}), all {all_count}, remaining {remaining_count}")
        if query is not None:


            item_data_all[category_id] = get_childs(query)


def get_categories_without_children(categories):
    result = {}
    for key, category in categories.items():
        childs = category.get("children")
        if childs:
            result = result | get_categories_without_children(childs)
        else:
            result[key]=category.get('query')

    return result


def save_nested_dict_to_excel(data: dict, filename: str):
    def write_sheet(ws, items):
        ws.append(["ID", "Название", "Ур. Вложенности","Бренд",'Варианты товара'])

        def write_childs(items):
            for key, item in items.items():
                item_list = [
                    key,
                    item["name"],
                    item["nesting_lv"],
                ]
                try:
                    brand = item["brand"]
                    if brand:
                        brand = f'\"{str(brand)}\"'
                except:
                    brand = ""
                item_list.append(brand)
                if item.get('colors'):
                    for color in item.get('colors'):
                        item_list.append(color['name'])
                ws.append(item_list)

        def write_row(item):
            ws.append([item["id"],item["name"], item["nesting_lv"]])
            children = item.get("children", {})
            if len(children) == 0:
                # get childs
                childs = item_data_all.get(item["id"], None)
                # add childs to page
                if childs:
                    write_childs(childs)
            for child in children.values():
                write_row(child)

        for item in items.values():
            write_row(item)

    wb = Workbook()
    wb.remove(wb.active)  # удалим дефолтный лист

    for sheet_name, content in data.items():

        children = content.get("children", None)
        query = content.get("query", None)
        if children is not None or query is not None:
            name = f"{content['name']}"
            ws = wb.create_sheet(title=name)
            write_sheet(ws, {sheet_name: content})

    wb.save(filename)


if __name__ == '__main__':
    all_categories = get_data_by_request(initial_address)
    categories = parse_childs(all_categories)

    categories_without_childs = get_categories_without_children(categories)

    load_childs(categories_without_childs)

    save_nested_dict_to_excel(categories, OUTPUT_FILE_NAME)