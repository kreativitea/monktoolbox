# -*- coding: utf-8 -*-
# Created by http://www.reddit.com/user/toshitalk
# Released under GPL.

# import standard libraries
import os
import sys
import json
import requests

from conn import download_data
from excel import write_data
from excel import select_excel
from calculations import get_data
from calculations import get_values


if __name__ == "__main__":
    data = download_data()

    try:
        tool_file = sys.argv[1]
    except IndexError:
        tool_file = select_excel()

    tool_path = os.path.join(os.getcwd(), tool_file)

    # currently only using stats.
    # actives and passives to be implemented later.
    stats, actives, passives = get_data(data)
    data = get_values(stats)
    write_data(tool_path, data)
