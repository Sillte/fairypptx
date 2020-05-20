import os
import re
from data_handler import listmap_handler

_this_folder = os.path.dirname(os.path.abspath(__file__))
_reference_folder = os.path.join(_this_folder, "object")
_instance_type_pattern = re.compile("<POINTER\((.+)\)")


def get_reference(typename):
    return listmap_handler.readcsv(
            os.path.join(_reference_folder,
                         typename.lower() + ".csv"))

def get_typename(pptx_object):
    s = str(pptx_object)
    s = s.replace("_", "")
    try:
        return _instance_type_pattern.findall(s)[0]
    except IndexError as e:
        pass
    return str(type(pptx_object))

def _pointer_converter(string):
    try:
        s = _instance_type_pattern.findall(string)[0]
    except:
        return string
    else:
        return "POINTER-" + s

    

def inspect(pptx_object, attr_type="property"):
    """ inspect the pptx_object.

    :param attr_type: ``method`` or ``property``. 

    :return: ``dict``.
    """
    attr_type = attr_type.lower()
    assert attr_type in ["property", "method"]
    try:
        s = str(pptx_object)
        original_typename = _instance_type_pattern.findall(s)[0]
        lower_typename = original_typename.lower()
        ref_listmap = get_reference(lower_typename)
    except Exception as e:
        raise e 
    
    result_dict = dict()

    for row in ref_listmap:
        try:
            if row["type"] != attr_type:
                continue
            name = row["name"]
            prop = str(getattr(pptx_object, row["name"]))
            result_dict[name] = _pointer_converter(prop)
        except Exception as e:
            #print(e)
            pass
            #print(e)
    return result_dict
            
    

if __name__ == "__main__":
    pass
