import win32gui
import win32con
from contextlib import contextmanager

## Public API

# ## [NOTICE] These variables requires changes in the future.  
# This `CLASSNAME` is different based on the version of `PPTX`.
# `_get_pptx_classname` is used for acquire..
PPTX_CLASSNAME = "PPTFrameClass"  
# This `CLASSNAME` is based on the debug, so I'm not so confident.  
MAIN_CLASSNAME = "MDIClient"  

@contextmanager
def lock_screen(invalidate=False):
    """Lock the Screen of Powerpoint inside the context. 

    Memorandum
    ----------
    I feel this is appropriate for drawing `Markdown`.  
    However, I feel effects of this is very little. 
    """

    main_hwnd = None
    try:
        if invalidate is True:
            raise RuntimeError("Invalid `False` is given.")  
        pptx_hwnd = get_pptx_hwnd()
        main_hwnd = to_main_hwnd(pptx_hwnd)
        win32gui.SendMessage(main_hwnd, win32con.WM_SETREDRAW, 0)
    except Exception as e:
        print("Not locking", type(e), e)
        is_succeeded = False
    else:
        is_succeeded = True

    def _recover():
        if main_hwnd is not None:
            win32gui.SendMessage(main_hwnd, win32con.WM_SETREDRAW, 1)

    try:
        yield
    except Exception as e:
        # No matter what, `WM_SETREDRAW` must be reversed.
        _recover()
        raise e
    else:
        _recover()


## Private API
##

def to_main_hwnd(pptx_hwnd):
    """Return Window Handle of the `Main Area` 
    Notice that mainly this is heavily relies on guessing. 

    It seems `MDOClient` is appripriate, but is it all o.k for
    all environment? 
    """

    hwnd = win32gui.FindWindowEx(pptx_hwnd, None, MAIN_CLASSNAME, None)
    if hwnd == 0:
        raise RuntimeError("Powerpoint `main` hwnd is not found.")
    return hwnd

def get_pptx_hwnd():
    from fairypptx import Application
    Application()
    hwnd = win32gui.FindWindow(PPTX_CLASSNAME, None)
    if hwnd == 0:
        print("`PPTX_CLASSNAME` is not valid.  ")
        print("Fallback is tried, but required change of implementation")
        classname = _get_pptx_classname()
        hwnd = win32gui.FindWindow(PPTX_CLASSNAME, None)
        if hwnd == 0:
            raise RuntimeError("Powerpoint hwnd is not found. ")
    assert hwnd != 0 
    return hwnd

def _to_special_hwnd():
    # May be other windowclass faster
    # So, here investigated, however,
    # I feel the speed does not change. 
    pptx_hwnd = get_pptx_hwnd()
    MAIN_CLASSNAME = "mdiClass"  
    result = None

    infos = []
    def callback(h, lp): 
        nonlocal result
        name = win32gui.GetClassName(h)
        if MAIN_CLASSNAME == name:
            result = h
            return False
        return True
    win32gui.EnumChildWindows(pptx_hwnd, callback, None)
    assert result is not None, "Result does not found"
    return result


def _get_pptx_classname():
    """ Based on guesses, acquire the `Window Class` of `PowerPoint`. 
    """
    from fairypptx import Application  # Hierarchy of dependency.
    Application()

    hwnd_to_text = dict()
    def callback(hwnd, lparam):
        text = win32gui.GetWindowText(hwnd)
        if text: 
            hwnd_to_text[hwnd] = text
    win32gui.EnumWindows(callback, None)
    targets = [hwnd for hwnd, text in hwnd_to_text.items() if text.endswith("PowerPoint")] 
    name = win32gui.GetClassName(targets[0])
    return name

if __name__ == "__main__":
    TEXT =  """ 
    ## SAMPLE
    """.strip()

    from fairypptx import Markdown
    import time

    start = time.time()
    with lock_screen(invalidate=False):
        Markdown.make(TEXT)
    locked_time = time.time() - start

    start = time.time()
    with lock_screen(invalidate=True):
        Markdown.make(TEXT)

    unlocked_time = time.time() - start

    print("Elapsed Time with locking.", locked_time)
    print("Elapsed Time without locking.", unlocked_time)
