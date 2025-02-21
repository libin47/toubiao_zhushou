import ttkbootstrap as ttk
from ttkbootstrap.tooltip import ToolTip


def tooltip(widget, text):
    """
    给组件添加提示
    """
    ToolTip(widget, text=text)
    return widget


def convert_to_json(obj):
    if isinstance(obj, ttk.StringVar):
        return obj.get()
    elif isinstance(obj, ttk.BooleanVar):
        return obj.get()
    elif isinstance(obj, ttk.IntVar):
        return obj.get()
    elif isinstance(obj, ttk.DoubleVar):
        return obj.get()
    elif isinstance(obj, dict):
        return {k: convert_to_json(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_json(item) for item in obj]
    elif isinstance(obj, str):
        return obj
    elif isinstance(obj, int):
        return obj
    elif isinstance(obj, float):
        return obj
    elif isinstance(obj, bool):
        return obj
    elif isinstance(obj, object):
        return convert_to_json(obj.__dict__)
    else:
        return obj

def json_to_obj(obj, json):
    if isinstance(json, dict):
        for k, v in json.items():
            if hasattr(obj, k):
                if isinstance(getattr(obj, k), ttk.StringVar):
                    getattr(obj, k).set(v)
                elif isinstance(getattr(obj, k), ttk.BooleanVar):
                    getattr(obj, k).set(v)
                elif isinstance(getattr(obj, k), ttk.IntVar):
                    getattr(obj, k).set(v)
                elif isinstance(getattr(obj, k), ttk.DoubleVar):
                    getattr(obj, k).set(v)
                elif isinstance(getattr(obj, k), dict):
                    json_to_obj(getattr(obj, k), v)
                elif isinstance(getattr(obj, k), list):
                    obj.__setattr__(k, v)
                else:
                    json_to_obj(getattr(obj, k), v)
                    # getattr(obj, k).set(v)
    else:
        obj.set(json)