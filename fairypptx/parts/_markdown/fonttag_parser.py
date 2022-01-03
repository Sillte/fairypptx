import re   
from ast import literal_eval
from fairypptx.color import Color

def parse_span_color_size(attributes):
    """ Interpret ``color`` and/or ``font-size`` element from <span> tag.
    """
    result_color_arg = None
    result_size_arg = None

    for attribute in attributes:
        a_key, a_value = attribute
        if a_key.lower() != "style":
            continue
        elements = a_value.split(";")
        print(elements)

        for element in elements:
            key, value = map(lambda s:s.strip().lower(), element.split(":"))
            if key == "font-size":
                cand = _size_arg_converter(value)
                if cand:
                    result_size_arg = cand
            if key == "color":
                cand = _color_arg_converter(value)
                if cand:
                    result_color_arg = cand

    return result_color_arg, result_size_arg


_args_pattern = re.compile(r'(".*")')
_color_pattern = re.compile("color=(.*)")
_size_pattern = re.compile("size=(.*)")

def parse_color_size(input_str):
    """ Interpret ``color`` and/or ``size`` element,
    from ``<font...> tag``.

    :return: If specified, ``Color`` and ``float`` are returned 
    respectively, For not-specified attribute, None is returned.
    """
    input_str = input_str[0:-1]
    tokens = list(_tokenize(input_str))

    result_color_arg = None
    result_size_arg = None

    for arg in tokens:
        color_ret = _color_pattern.findall(arg)
        if color_ret: 
            cand = _color_arg_converter(color_ret[0])
            if cand: 
                result_color_arg = cand

        size_ret = _size_pattern.findall(arg)
        if size_ret: 
            cand = _size_arg_converter(size_ret[0])
            if cand:
                result_size_arg = cand

    return result_color_arg, result_size_arg


def _tokenize(input_str):
    """ This is an imcomplete tokenizer,
    There is no consideration as for co-existence of " and '.
    """
    _quotation_parity = 0
    buf = ""
    for s in input_str:
        if s == " " and (_quotation_parity % 2 == 0):
            yield buf
            buf = ""
            continue
        if s in {"'", '"'}:
            _quotation_parity += 1
        buf += s
    yield buf


def _size_arg_converter(size_arg):
    """ If it succeeds to interpret, then float is returned,
    otherwise, ``None`` is returned.
    """
    size_arg = size_arg.replace('"', '')
    size_arg = size_arg.replace("'", '')
    size_arg = size_arg.lower().replace("px", "")
    size_arg = size_arg.lower().replace("pt", "")
    try:
        return float(size_arg)
    except Exception as e:
        print("Exception", e, type(e), size_arg)
    return None

def _color_arg_converter(color_arg):
    """ If it succeeds to interpret, then ``Color`` is returned,
    otherwise, ``None`` is returned.
    """
    color_arg = color_arg.replace('"', '')
    color_arg = color_arg.replace("'", '')
    if color_arg.lower().find("rgb") != -1:
        color_arg = color_arg.lower().replace("rgb", "")
        color_arg = literal_eval(color_arg) # Evaluation of tuple.
    try:
        return Color(color_arg)
    except Exception as e:
        print("Exception", e, type(e), color_arg)
    return None

if __name__ == "__main__":
    color = Color("#FF38FF")
    print(color.as_hex())
    input_str = '<font color="rgb(255,0, 0)" size="14pt">'
    color_arg, size_arg = parse_color_size(input_str)
    print(color_arg, size_arg)

