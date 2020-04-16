from collections import OrderedDict
import datetime
import io
import os
import time

from anki.lang import _
from aqt import mw
from aqt.utils import showInfo
from aqt.qt import *


from .config import gc


def due_day(card):
    if card.queue <= 0:
        return ""
    else:
        if card.queue in (2, 3):
            if card.odue:
                myvalue = card.odue
            else:
                myvalue = card.due
            mydue = time.time()+((myvalue - mw.col.sched.today)*86400)
        else:
            if card.odue:
                mydue = card.odue
            else:
                mydue = card.due
        return time.strftime("%Y-%m-%d", time.localtime(mydue))


def valueForOverdue(card):
    return mw.col.sched._daysLate(card)


def percent_overdue(card):
    overdue = mw.col.sched._daysLate(card)
    ivl = card.ivl
    if ivl > 0:
        return "{0:.2f}".format((overdue+ivl)/ivl*100)


def fmt_long_string(name, value):
    l = 0
    u = value
    out = ""
    while l < len(name):
        out += name[l:l+u] + '\n'
        l += u
    return out.rstrip('\n')


def allRevsForCard(cid):
    entries = mw.col.db.all(
        "select id/1000.0, ease, ivl, factor, time/1000.0, type "
        "from revlog where cid = ?", cid)
    if not entries:
        return ""
    allRevs = ""
    for (date, ease, ivl, factor, taken, type) in entries:
        tstr = [_("Lrn"), _("Rev"), _("ReLn"), _("Filt"), _("Resch")][type]
            # Learned, Review, Relearned, Filtered, Defered (Rescheduled)
        int_due = "na"
        if ivl > 0:
            int_due_date = time.localtime(date + (ivl * 24 * 60 * 60))
            int_due = time.strftime(_("%Y-%m-%d"), int_due_date)
        allRevs += "#".join((time.strftime("%Y-%m-%dT@%H:%M", time.localtime(date)), 
                             str(tstr),
                             str(ease),
                             str(ivl),
                             str(int_due),
                             str(int(factor / 10)) if factor else "")) + '-----'
    return allRevs


def getSaveDir(parent, title, identifier_for_last_user_selection):
    config_key = identifier_for_last_user_selection + 'Directory'
    defaultPath = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
    path = mw.pm.profile.get(config_key, defaultPath)
    dir = QFileDialog.getExistingDirectory(parent, title, path, QFileDialog.ShowDirsOnly)
    return dir


def now():    #time
    CurrentDT=datetime.datetime.now()
    return CurrentDT.strftime("%Y-%m-%dT%H-%M-%S")