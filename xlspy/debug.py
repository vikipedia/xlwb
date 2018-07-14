from functools import wraps
import os

def debug(f):
    depth  = 0
    
    @wraps(f)
    def wrapper(*args, **kwargs):
        nonlocal depth
        depth += 1
        print(f.__qualname__, depth, args, kwargs)
        v = f(*args, **kwargs)
        depth -= 1
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
        r = f(*args, **kwargs)
        log("| "*level + "|--" +  "return", r)
        level -= 1
        return r
    return wrapper
        


def debugmethods(cls):

    for k,v in vars(cls).items():
        if callable(v):
            setattr(cls, k, debug(v))
    return cls
