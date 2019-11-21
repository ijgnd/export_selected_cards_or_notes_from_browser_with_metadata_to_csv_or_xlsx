from aqt import mw


def gc(arg, fail=False):
    # some TODOS ...
    if arg == "remove_newlines":
        return True
    else:
        return mw.addonManager.getConfig(__name__).get(arg, fail)