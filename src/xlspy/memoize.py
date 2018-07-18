from functools import wraps

def memoize(f):
    cache = {}

    @wraps(f)
    def wrapper(*args):
        if args not in cache:
            cache[args] = f(*args)
        return cache[args]

    
    return wrapper
