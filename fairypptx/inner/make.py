import os, glob, shutil
import itertools
import lxml.html
import requests
from data_handler import listmap_handler

_main_host = "https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles"
_this_folder = os.path.dirname(os.path.abspath(__file__))

def _create_folder(folder):
    if os.path.exists(folder) is False:
        os.mkdir(folder)

"""
def _shape_to_address(object_instance):
    def _to_str(object_instance):
        if not isinstance(object_instance, str):
            return object_instance.name
        return object_instance
        
    address = _main_host + "/{}-object-powerpoint" \
                           .format(_to_str(object_instance))
    return address

def _debug_write(filename):
    with open(filename, "w", encoding="utf8") as fp:
        fp.write(text)
"""

def _get_objects():
    output_folder = "object"
    _create_folder(os.path.join(_this_folder, output_folder))
    address = _main_host + "/object-model-powerpoint-vba-reference"
    res = requests.get(address)
    text = res.text

    def _get_object_links(links):
        listmap = list()
        for link in links:
            try:
                href = link.get("href")
                last_path = href.split("/")[-1]
                identifiers = last_path.split("-")
                name = identifiers[0]
                assert identifiers[1] == "object"
            except Exception as e:
                print(e)
            else:
                listmap.append({"name":name, "href":href, "link":link})
        return listmap

    def _to_prop_row(href):
        last_path = href.split("/")[-1]
        identifiers = last_path.split("-")
        name = identifiers[-3]
        attr_type = identifiers[-2]
        return {"name":name, "type":attr_type}

    root = lxml.html.fromstring(text)
    links = root.xpath("//a")
    object_link_listmap = _get_object_links(links)
    object_links = [row["link"] for row in object_link_listmap]

    for index, link in enumerate(object_links):
        name = object_link_listmap[index]["name"]
        elem = link.getparent().getnext()
        targets = elem.xpath(".//a")
        listmap = list()
        for t in targets:
            row = dict()
            row["href"] = t.get("href")
            row.update(_to_prop_row(row["href"]))
            listmap.append(row)
        listmap_handler.writecsv(os.path.join(output_folder, name + ".csv"), listmap)
    ref_listmap = [{"name":row["name"], "href":row["href"]} for row in listmap]
    listmap_handler.writecsv(os.path.join(_this_folder, output_folder + ".csv"), listmap)
    return ref_listmap


if __name__ == "__main__":
    listmap = _get_objects();
    listmap_handler.writecsv("objects.csv", listmap)

