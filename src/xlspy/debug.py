from functools import wraps
import os

def debug(f):
    @wraps(f)
    def wrapper(*args, **kwargs):

        print(f.__qualname__,  args, kwargs)
        v = f(*args, **kwargs)
        return v

    return wrapper



level = 0

def log(*args):
    if os.getenv("DEBUG")=="true":
        print(*args)
        
def trace(f):
    
    def wrapper(*args, **kwargs):
        global level
        log("| "*level + "|--" + f.__name__ ,  *args)
        level += 1
        r = None
        try:
            r = f(*args, **kwargs)
        finally:
            log("| "*level + "|--" +  "return", r)
            level -= 1
        return r
    return wrapper
        


def debugmethods(cls):

    for k,v in vars(cls).items():
        if callable(v):
            setattr(cls, k, debug(v))
    return cls
