import requests


def download_excel(url, path):
    data = requests.get(url).content
    f = open(path, "wb")
    f.write(data)
    f.close()

