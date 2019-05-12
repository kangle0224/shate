import json
from codecs import open
from datetime import datetime, date, timedelta


def write_json(filename, data):
    with open(filename, "w", encoding="utf-8") as f:
        f.write(json.dumps(data, indent=2))
