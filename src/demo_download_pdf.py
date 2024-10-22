import os

import requests


def download_pdf(url, save_path):
    response = requests.get(url)
    if response.status_code == 200:
        with open(save_path, "wb") as f:
            f.write(response.content)
        print(f"PDF downloaded successfully and saved to {save_path}")
    else:
        print(f"Failed to download PDF. Status code: {response.status_code}")


if __name__ == "__main__":
    file_path = "D:/projects/pdf2excel/pdf2excel_AD/others/AD_PO.txt"

    with open(file_path, "r", encoding="utf8") as file:
        org_url = file.read()

    # print(f"org_url:{org_url}")

    org_url_list = org_url.split()

    # print(f"org url list:{org_url_list}")
    head_url = "http://192.168.0.120/pic"
    full_url_list = [head_url + tmp_org_url for tmp_org_url in org_url_list]

    # print(f"full_url_list:{full_url_list}")

    out_path = "D:/projects/pdf2excel/pdf2excel_AD/others/samples"

    for tmp_full_url in full_url_list:
        file_name = tmp_full_url.split("/")[-1]
        # print(f"file_name:{file_name}")
        download_pdf(url=tmp_full_url, save_path=os.path.join(out_path, file_name))
