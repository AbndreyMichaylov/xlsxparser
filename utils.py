

def int_try_parse(object, default=None):
    try:
        return int(object)
    except:
        return default