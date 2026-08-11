from datetime import datetime
from zoneinfo import ZoneInfo


def create_file_name(first_name):
    now = datetime.now(ZoneInfo("Asia/Kathmandu"))
    month = now.strftime("%B")
    day = now.day
    print(day)

    return f"{first_name}Transfer_{month}{day}.xlsx"


if __name__ == "__main__":
    print(create_file_name("Baneshwor"))
