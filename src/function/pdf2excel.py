import re
import sys
from typing import List

import json

import pandas as pd
import pymupdf

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def clean_annot_in_doc(doc):
    # remove annotation information from pdf files
    # to avoid the impact of annotation information on form extraction
    for page in doc:
        for annot in page.annots():
            page.delete_annot(annot=annot)


def sort_size_list(size_set):
    size_to_number = {
        "XXXS": 1,
        "XXS": 2,
        "XS": 3,
        "S": 4,
        "M": 5,
        "L": 6,
        "XL": 7,
        "XXL": 8,
        "XXXL": 9,
    }

    def sort_sizes_str(sizes):
        sorted_sizes = sorted(
            sizes, key=lambda size: size_to_number.get(size, 0), reverse=False
        )
        return sorted_sizes

    bool_all_number = True
    size_list = []
    for tmp_size_set in size_set:
        try:
            # tmp_size_set = int(tmp_size_set)
            size_list.append(int(tmp_size_set))
        except Exception as e:
            size_list.append(tmp_size_set)
            bool_all_number = False

    # for tmp_size in size_list:
    #     print(f"tmp_size:{tmp_size}")
    # print(f"bool_all_number:{bool_all_number}")
    if bool_all_number is False:
        sorted_size = sort_sizes_str(size_list)
    else:
        sorted_size = sorted(size_list, reverse=False)

    return sorted_size


def trans_json2ws(total_style_info_list):
    new_df = pd.DataFrame(total_style_info_list)
    # new_df.rename(
    #     columns={
    #         "Reference": "Reference",
    #         "COLOUR": "颜色",
    #         "Ref. AD": "款号",
    #         "Price": "Price",
    #         "Delivery\ndate": "ETD",
    #     },
    #     inplace=True,
    # )

    print(f"new_df:{new_df}")

    # new_df["单价"] = new_df["单价"].astype(float)
    # new_df["数量"] = new_df["数量"].astype(float)
    # new_df["总金额"] = new_df["总金额"].astype(float)

    # front_list = ["客户", "季度", "款号", "颜色", "Reference", "PO", "国家"]
    # end_list = ["数量", "单价", "总金额", "离厂时间"]

    # size_list = sort_size_list(size_set=size_columns_set)

    # print(f"size_list:{size_list}")
    # new_order = front_list + size_list + end_list
    # exist_index = new_df.columns
    # print(f"new_df.columns:{new_df.columns}")
    # for ordered_key in new_order:
    #     if ordered_key in exist_index:
    #         pass
    #     else:
    #         new_df[ordered_key] = pd.NA
    # new_df = new_df[new_order]

    # if "DESCRIPTION" in new_df.columns:
    #     new_df.drop(columns="DESCRIPTION", inplace=True)

    # print(f"new_df:{new_df}")

    wb = Workbook()

    # 获取当前活跃的工作表
    ws = wb.active

    # 将DataFrame的数据写入工作表
    for r in dataframe_to_rows(new_df, index=False, header=True):
        print(f"r:{r}")
        ws.append(r)
    return wb


def func_pdf2excel(pdf_content):
    """transform pdf to excel

    Args:
        pdf_content (Union[str, bytes]): file path or bytes
        size_columns_set (set, optional): the set of size columns. Defaults to set().

    Returns:
        _type_: _description_
    """
    if isinstance(pdf_content, str):
        doc = pymupdf.open(pdf_content)
    else:
        doc = pymupdf.open(stream=pdf_content)

    clean_annot_in_doc(doc=doc)

    tar_content_list = ["Version:", ".com"]
    search_content_res_list = find_target_block_content(
        page=doc[0], tar_content_list=tar_content_list
    )

    print(f"search_content_res_list:{search_content_res_list}")

    PO = search_content_res_list[0].split()[0].strip()
    Season = search_content_res_list[1].split()[-2].strip()

    # find country

    Country = ""

    page = doc[0]
    page_content = page.get_text(option="dict")
    delivery_info = {"pos": (), "content": ""}
    for block in page_content["blocks"]:
        tmp_block_list = []
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    tmp_block_list.append(span["text"])
                    # if "Delivery" in span["text"]:
                    #     delivery_info["content"] = span["text"]
                    #     delivery_info["pos"] = span["bbox"]
                    #     break
        block_content = " ".join(tmp_block_list)
        if block_content.strip() == "Delivery":
            delivery_info["content"] = span["text"]
            delivery_info["pos"] = span["bbox"]
            break
    print(f"delivery_info:{delivery_info}")
    delivery_x = delivery_info["pos"][0]
    delivery_y = delivery_info["pos"][3]

    page_content = page.get_text(option="dict")
    target_block = {"pos": (), "content": ""}
    for block in page_content["blocks"]:
        if "lines" in block:
            tmp_block_pos = block["bbox"]
            tmp_block_list = []
            block_x = tmp_block_pos[0]
            block_y = tmp_block_pos[1]
            # print(f"tmp block:{block['bbox']}")

            if (
                block_x > delivery_x - delivery_x * 0.1
                and block_x < delivery_x + delivery_x * 0.1
                and block_y > delivery_y
            ):
                print("get block")
                if target_block["pos"]:
                    if block_y < target_block["pos"][1]:
                        target_block["pos"] = tmp_block_pos
                    else:
                        continue
                else:
                    target_block["pos"] = tmp_block_pos

                for line in block["lines"]:
                    for span in line["spans"]:
                        tmp_block_list.append(span["text"])

                tmp_block_content = " ".join(tmp_block_list)
                target_block["content"] = tmp_block_content
                # print(f'content:{target_block["content"]}')
        # total_block_content_list.append(tmp_block_content)
    # print(f"target_block:{target_block}")
    Country = target_block["content"].split(" ")[-1].strip()
    print(f"PO:{PO} Season:{Season} Country:{Country}")
    # size_columns_set = set()
    # for key, val in cols_size.items():
    #     size_columns_set.add(val)

    page = doc[0]
    tables = page.find_tables()
    print("get tables")
    for table in tables:
        table_df = table.to_pandas()
        table_json = json.loads(table_df.to_json(orient="records", force_ascii=False))
        print(f"table_json:{table_json}")

        new_table_json = []
        for tmp_table_row in table_json:
            new_row_info = {"客户": "AD", "季度": Season, "国家": Country}
            for key, value in tmp_table_row.items():
                if value:
                    if "Col" in key and key != "Colour":
                        key = "Quantity"
                    try:
                        key = int(key)
                        value = int(value)
                        size_columns_set.add(key)
                        # new_row_info['尺码']
                    except Exception as e:
                        pass
                    new_row_info[key] = value
            new_table_json.append(new_row_info)

        print(f"new_table_json:{new_table_json}")

        break

    table_json = new_table_json
    new_row_info_list = []
    # transform org table to target table
    for tmp_table_row in table_json:
        # print(f'org price:{tmp_table_row["Price"]}')
        price_list = tmp_table_row["Price"].split(",")
        # print(f"price_list:{price_list}")
        # print(f"len :{len(price_list[1].strip())}")
        price = int(price_list[0].strip()) + 0.1 ** len(price_list[1].strip()) * int(
            price_list[1].strip()
        )
        # print(f"price:{price}")

        for key, value in tmp_table_row.items():
            tmp_new_row_info = {
                "客户": tmp_table_row["客户"],
                "季度": Season,
                "款号": tmp_table_row["Ref. AD"],
                "颜色": tmp_table_row["Colour"],
                "Reference": tmp_table_row["Reference"],
                "PO": PO,
                "国家": Country,
                "尺码": 0,
                "PO数量": 0,
                "ERP数量": 0,
                "辅料数量": 0,
                "ETD": tmp_table_row["Delivery\ndate"],
                "Price": price,
                "Amount": 0,
            }
            if isinstance(key, int):
                tmp_new_row_info["尺码"] = key
                tmp_new_row_info["PO数量"] = value
                if tmp_new_row_info["国家"] == "Mexico":
                    # 非国际单
                    tmp_new_row_info["国家"] = "墨西哥单"
                    tmp_new_row_info["ERP数量"] = value
                else:
                    # 国际单
                    tmp_new_row_info["国家"] = "国际单"
                    tmp_new_row_info["ERP数量"] = round_up(value * 1.03)

                tmp_new_row_info["辅料数量"] = round_up(
                    tmp_new_row_info["ERP数量"] * 1.06
                )
                tmp_new_row_info["Amount"] = (
                    tmp_new_row_info["ERP数量"] * tmp_new_row_info["Price"]
                )
                new_row_info_list.append(tmp_new_row_info)

    if isinstance(pdf_content, str):
        wb = trans_json2ws(
            total_style_info_list=new_row_info_list,
            # size_columns_set=size_columns_set,
        )
        wb.save("a.xlsx")
    else:
        return new_row_info_list


def mark_pdf(input_path: str, output_path: str = "./", level: str = ""):
    """mark PDF content with rectangle

    Args:
        input_path (str): input file path
        output_path (str, optional): output file folder path. Defaults to "./".
        level (str, optional): mark content tags, include block, line, span, table, cell. Defaults to "".
    """
    doc = pymupdf.open(input_path)

    for page in doc:
        page_content = page.get_text(option="dict")

        for block in page_content["blocks"]:
            if level == "block":
                page.draw_rect(pymupdf.Rect(block["bbox"]), color=(1, 0, 0))

            if "lines" in block:
                for line in block["lines"]:
                    if level == "line":
                        page.draw_rect(pymupdf.Rect(line["bbox"]))
                    for span in line["spans"]:
                        if level == "span":
                            page.draw_rect(pymupdf.Rect(span["bbox"]))

        if level in ["table", "cell"]:
            tables = page.find_tables()
            print("get tables")

            if level == "cell":
                for table in tables.tables:
                    for cell in table.cells:
                        table.page.draw_rect(cell, color=(1, 0, 0))

            if level == "table":
                for table in tables.tables:
                    table.page.draw_rect(table.bbox, color=(0, 1, 0))

    doc.save(f"{output_path}/{level}.pdf")


def get_page_content(page) -> List[str]:
    """get page content in PDF

    Args:
        page (_type_): page object

    Returns:
        List[str]: content list
    """

    page_content = page.get_text(option="dict")
    total_block_content_list = []
    for block in page_content["blocks"]:
        tmp_block_content = ""
        tmp_block_list = []
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    tmp_block_list.append(span["text"])

        tmp_block_content = " ".join(tmp_block_list)
        total_block_content_list.append(tmp_block_content)

    return total_block_content_list


def find_target_block_content(page, tar_content_list: List[str]) -> List[str]:
    """get target block content include tar_content_list

    Args:
        tar_content_list (List[str]): _description_

    Returns:
        List[str]: _description_
    """
    # tar_content_list = ["camilla and marc Order No:", "EX FACTORY DATE:"]
    search_content_res_list = []
    # page = doc[0]
    page_content_list = get_page_content(page=page)
    print(f"page_content_list:{page_content_list}")
    for tmp_tar_content_list in tar_content_list:
        bool_find = False
        for tmp_page_content_list in page_content_list:
            if tmp_tar_content_list in tmp_page_content_list:
                search_content_res_list.append(tmp_page_content_list)
                bool_find = True
                break

        if bool_find is False:
            search_content_res_list.append("")

    return search_content_res_list


def round_up(num):
    """数值向上取整

    Args:
        num (float): 数值

    Returns:
        int: 向上取整后的数值
    """
    return -(-num // 1)


if __name__ == "__main__":
    ORG_PDF_PATH = (
        "D:/projects/pdf2excel/pdf2excel_AD/others/sample_file/2.17.16.2499国际单.pdf"
    )

    # mark_pdf(input_path=ORG_PDF_PATH, output_path="./2499", level="cell")
    func_pdf2excel(pdf_content=ORG_PDF_PATH)
    # pdf_content = ORG_PDF_PATH
    # if isinstance(pdf_content, str):
    #     doc = pymupdf.open(pdf_content)
    # else:
    #     doc = pymupdf.open(stream=pdf_content)

    # clean_annot_in_doc(doc=doc)

    # tar_content_list = ["Version:", ".com"]
    # search_content_res_list = find_target_block_content(
    #     tar_content_list=tar_content_list
    # )

    # print(f"search_content_res_list:{search_content_res_list}")

    # PO = search_content_res_list[0].split()[0].strip()
    # Season = search_content_res_list[1].split()[-2].strip()

    # # find country

    # Country = ""

    # page = doc[0]
    # page_content = page.get_text(option="dict")
    # delivery_info = {"pos": (), "content": ""}
    # for block in page_content["blocks"]:
    #     tmp_block_list = []
    #     if "lines" in block:
    #         for line in block["lines"]:
    #             for span in line["spans"]:
    #                 tmp_block_list.append(span["text"])
    #                 # if "Delivery" in span["text"]:
    #                 #     delivery_info["content"] = span["text"]
    #                 #     delivery_info["pos"] = span["bbox"]
    #                 #     break
    #     block_content = " ".join(tmp_block_list)
    #     if block_content.strip() == "Delivery":
    #         delivery_info["content"] = span["text"]
    #         delivery_info["pos"] = span["bbox"]
    #         break
    # print(f"delivery_info:{delivery_info}")
    # delivery_x = delivery_info["pos"][0]
    # delivery_y = delivery_info["pos"][3]

    # page_content = page.get_text(option="dict")
    # target_block = {"pos": (), "content": ""}
    # for block in page_content["blocks"]:
    #     if "lines" in block:
    #         tmp_block_pos = block["bbox"]
    #         tmp_block_list = []
    #         block_x = tmp_block_pos[0]
    #         block_y = tmp_block_pos[1]
    #         # print(f"tmp block:{block['bbox']}")

    #         if (
    #             block_x > delivery_x - delivery_x * 0.1
    #             and block_x < delivery_x + delivery_x * 0.1
    #             and block_y > delivery_y
    #         ):
    #             print("get block")
    #             if target_block["pos"]:
    #                 if block_y < target_block["pos"][1]:
    #                     target_block["pos"] = tmp_block_pos
    #                 else:
    #                     continue
    #             else:
    #                 target_block["pos"] = tmp_block_pos

    #             for line in block["lines"]:
    #                 for span in line["spans"]:
    #                     tmp_block_list.append(span["text"])

    #             tmp_block_content = " ".join(tmp_block_list)
    #             target_block["content"] = tmp_block_content
    #             # print(f'content:{target_block["content"]}')
    #     # total_block_content_list.append(tmp_block_content)
    # # print(f"target_block:{target_block}")
    # Country = target_block["content"].split(" ")[-1].strip()
    # print(f"PO:{PO} Season:{Season} Country:{Country}")
    # # size_columns_set = set()
    # # for key, val in cols_size.items():
    # #     size_columns_set.add(val)

    # page = doc[0]
    # tables = page.find_tables()
    # print("get tables")
    # for table in tables:
    #     table_df = table.to_pandas()
    #     table_json = json.loads(table_df.to_json(orient="records", force_ascii=False))
    #     print(f"table_json:{table_json}")

    #     new_table_json = []
    #     for tmp_table_row in table_json:
    #         new_row_info = {"客户": "AD", "季度": Season, "国家": Country}
    #         for key, value in tmp_table_row.items():
    #             if value:
    #                 if "Col" in key and key != "Colour":
    #                     key = "Quantity"
    #                 try:
    #                     key = int(key)
    #                     value = int(value)
    #                     size_columns_set.add(key)
    #                     # new_row_info['尺码']
    #                 except Exception as e:
    #                     pass
    #                 new_row_info[key] = value
    #         new_table_json.append(new_row_info)

    #     print(f"new_table_json:{new_table_json}")

    #     break

    # table_json = new_table_json
    # new_row_info_list = []
    # # transform org table to target table
    # for tmp_table_row in table_json:
    #     # print(f'org price:{tmp_table_row["Price"]}')
    #     price_list = tmp_table_row["Price"].split(",")
    #     # print(f"price_list:{price_list}")
    #     # print(f"len :{len(price_list[1].strip())}")
    #     price = int(price_list[0].strip()) + 0.1 ** len(price_list[1].strip()) * int(
    #         price_list[1].strip()
    #     )
    #     # print(f"price:{price}")

    #     for key, value in tmp_table_row.items():
    #         tmp_new_row_info = {
    #             "客户": tmp_table_row["客户"],
    #             "季度": Season,
    #             "款号": tmp_table_row["Ref. AD"],
    #             "颜色": tmp_table_row["Colour"],
    #             "Reference": tmp_table_row["Reference"],
    #             "PO": PO,
    #             "国家": Country,
    #             "尺码": 0,
    #             "PO数量": 0,
    #             "ERP数量": 0,
    #             "辅料数量": 0,
    #             "ETD": tmp_table_row["Delivery\ndate"],
    #             "Price": price,
    #             "Amount": 0,
    #         }
    #         if isinstance(key, int):
    #             tmp_new_row_info["尺码"] = key
    #             tmp_new_row_info["PO数量"] = value
    #             if tmp_new_row_info["国家"] == "Mexico":
    #                 # 非国际单
    #                 tmp_new_row_info["ERP数量"] = value
    #             else:
    #                 # 国际单
    #                 tmp_new_row_info["ERP数量"] = round_up(value * 1.03)

    #             tmp_new_row_info["辅料数量"] = round_up(
    #                 tmp_new_row_info["ERP数量"] * 1.06
    #             )
    #             tmp_new_row_info["Amount"] = (
    #                 tmp_new_row_info["ERP数量"] * tmp_new_row_info["Price"]
    #             )
    #             new_row_info_list.append(tmp_new_row_info)

    # if isinstance(pdf_content, str):
    #     wb = trans_json2ws(
    #         total_style_info_list=new_row_info_list,
    #         # size_columns_set=size_columns_set,
    #     )
    #     wb.save("a.xlsx")
    # else:
    #     return new_row_info_list
    #     # new_df.to_excel("a.xlsx")

    # total_style_info_list, size_columns_set = func_pdf2excel(pdf_content=ORG_PDF_PATH)
    # pdf_content = ORG_PDF_PATH
    # func_pdf2excel(pdf_content=pdf_content)
