from functools import wraps, partialmethod
from django.views.decorators.csrf import csrf_exempt

def login_exempt(view_func):
    @csrf_exempt
    @wraps(view_func, updated=())
    def _wrapped_view(request, *args, **kwargs):
        setattr(request, '_login_exempt', True)
        return view_func(request, *args, **kwargs)

    return partialmethod(_wrapped_view, __wrapped__=view_func)

