from functools import wraps


def add_method(cls: object):
    """
    Decorator method to dynamically add a function to a particular class
    :param cls: class to add the decorated method to
    :return: the decorated method
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        setattr(cls, func.__name__, wrapper)
        return func
    return decorator
